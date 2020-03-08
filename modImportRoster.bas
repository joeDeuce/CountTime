Attribute VB_Name = "modImportRoster"
Option Compare Database
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'used to make sure we don't show file dialog
'after file has been selected
Global PDFSelected As Boolean

Sub SleepVBA(Millies As Long)
    
    Sleep (Millies)

End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
    
    FileExists = (Dir(FileToTest) <> "")

End Function

Public Function ReadImportedTextfile(ByRef f As Form) As Boolean
    
    f.bInline.Visible = True
    f.bInline.Width = 0
    f.bOutline.Visible = True
    f.lblLoad.Caption = "Reading alpha roster contents..."
    f.lblLoad.Visible = True
    
    Dim intFile As Integer
    Dim strFile As String
    Dim strIn As String
    Dim strOut As String
    Dim CleanString As Boolean
    Dim FirstLine As String
    Dim SecondLine As String
    Dim WrongPrison As Variant
    Dim WrongPrisonAdd As String
    
    CleanString = False
    strOut = vbNullString
    intFile = FreeFile()
    strFile = GetDBPath & "import.txt"
    
    If strFile <> "import.txt" Then
        Open strFile For Input As #intFile

        Set db = CurrentDb()
        
        ''wait until AFTER roster sanity check before
        ''we clean up last import table
        'db.Execute "DELETE * FROM [_tRosterImport]"
        
        Dim i, j, k As Integer
        i = 1
        j = 0
        k = 0
        'check first lines to make sure it is an alpha roster
        Line Input #intFile, FirstLine
        Line Input #intFile, SecondLine
        'Debug.Print FirstLine
        'Debug.Print SecondLine
        If (InStr(1, FirstLine, "WALKER") = 0) Or (InStr(1, SecondLine, "Alpha") = 0) Then
        
            'lets get the name of the wrong prison
            If (InStr(1, SecondLine, "Alpha") = 1) Then
            
                WrongPrison = Split(FirstLine, "               ")
                'Debug.Print Trim(WrongPrison(1))
                WrongPrisonAdd = "This roster is from " & Trim(WrongPrison(1)) & "!" & vbNewLine
                
            Else
            
                WrongPrisonAdd = "Alpha roster in unrecognized format." & vbNewLine
                
            End If
            
            MsgBox WrongPrisonAdd & "Make sure you are importing an alpha roster from Walker." & vbNewLine & vbNewLine & "Unable to import!", vbExclamation, "Error"
            Close #intFile
            ReadImportedTextfile = False
            
            f.Option1.Enabled = True
            f.OptionLabel1.Enabled = True
            
            RemoveImportTextfile

            Exit Function
        
        End If
        
        'clean the temp import table
        db.Execute "DELETE * FROM [_tRosterImport]"
        
        Do While Not EOF(intFile)
            
            Line Input #intFile, strIn
            
            CleanString = ExtractViaRegExp(strIn, db)
            i = i + 1
            
            'used to display a loading meter
            j = i * 9.5
            If j > f.bOutline.Width Then j = f.bOutline.Width
            k = j / 72
            If k > 100 Then k = 100
            f.bInline.Width = j
            f.lblLoad.Caption = "Importing inmates... " & k & "% complete"
            f.Option1.Enabled = False
            f.OptionLabel1.Enabled = False

            ''this sometimes causes things to happen out of order
            ''must make sure checks are made in other places to prevent
            ''erratic behavior
            'SleepVBA 1
            'DoEvents
            If i Mod 3 = 0 Then DoEvents
            
        Loop
        
        Close #intFile
        
        ReadImportedTextfile = True
        
    Else
        
        ReadImportedTextfile = False
        f.Option1.Enabled = True
        f.OptionLabel1.Enabled = True

    End If

End Function

Public Function ExtractViaRegExp(inpStr, ByRef dbp)
    
    Dim midStr
    Dim oRe, oMatch As Match, oMatches As MatchCollection
    Set oRe = New RegExp
    
    ' replace ISSG with 13
    midStr = Replace(inpStr, " ISSG", " 13")
    
    ' pattern to retrive fields from roster saved as text
    
    'original
    oRe.Pattern = "([A-Z\-]+), ([A-Z]+[ ][A-Z]*) +[A-Z \-,.]*([\d]+) *\w*[A-Z \-,./\d]* ([\d]+)-[NSEW]-([\d]+)"
    
    'updated thanks to orange from utteraccess
    'after testing there is no benefit over original
    'oRe.Pattern = "([A-Z\-]+),\s([A-Z]+\s?[A-Z]+?)\s+(\d{1,11})\s+.+\b(\d{1,2})-[NSEW]-(\d{1,2})"
    
    ' Get the Matches collection
    Set oMatches = oRe.Execute(midStr)
    
    'nothing to load from alpha roster
    If oMatches.Count = 0 Then Exit Function
    
    ' Get the first item in the Matches collection
    Set oMatch = oMatches(0)
    
    'Debug.Print oMatch.SubMatches(2)
    
    If oMatch.SubMatches(0) <> "" Then
        dbp.Execute "INSERT INTO [_tRosterImport] (LastName, FirstName, GDCNum, Dorm, Bed) VALUES ('" & oMatch.SubMatches(0) & "', '" & RTrim(oMatch.SubMatches(1)) & "', " & oMatch.SubMatches(2) & ", " & oMatch.SubMatches(3) & ", " & oMatch.SubMatches(4) & ")"
    End If
    
    ExtractViaRegExp = True

End Function

Public Function SplitRe(Text As String, Pattern As String, Optional IgnoreCase As Boolean) As String()
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.MultiLine = True
    End If

    re.IgnoreCase = IgnoreCase
    re.Pattern = Pattern
    SplitRe = Strings.Split(re.Replace(Text, ChrW(-1)), ChrW(-1))
End Function

Public Function ImportRoster()
    Dim IsRead As Boolean
    Dim UpdatedInmates As Integer
    Dim NewInmates As Integer
    Dim NewTIC, OldTIC As Integer
    Dim Switch As Form
    Set Switch = Forms!Switchboard
    
    IsRead = ReadImportedTextfile(Switch)

    Switch.lblLoad.Caption = "Import Complete..."
    Switch.bInline.Width = Switch.bOutline.Width
    Switch.Option1.Enabled = True
    Switch.OptionLabel1.Enabled = True
        
    If IsRead = True Then
        'after importing the roster, we need to update DormAtCount on all open counts
        
        'TODO: Option to NOT sanitize open counts, as this can delete a complete count
        'if Alpha Roster import goes awry... UPDATE: new option not needed, with new
        'sanity checks on import
        
        OldTIC = DCount("GDCNum", "tBaseline", "InActive = 0")
        
        Set db = CurrentDb()
        db.Execute "1qUpdateToInactive"
        db.Execute "2qUpdateExisting"
        UpdatedInmates = db.RecordsAffected
        db.Execute "22qAppendNew"
        NewInmates = db.RecordsAffected
        db.Execute "3qUpdateToActive"
        NewTIC = db.RecordsAffected
        db.Execute "4qUpdateOpenCounts"
        'TODO need to look at alternatives for sanitize
        db.Execute "5qSanitizeOpenCounts"

        MsgBox "Successfully imported! " & vbNewLine & vbNewLine & NewInmates & " new inmates. " & vbNewLine & UpdatedInmates & " existing or returning inmates updated or verified. " & vbNewLine & vbNewLine & "Old TIC: " & OldTIC & vbNewLine & "New TIC: " & NewTIC, vbInformation, "Success!"
        SleepVBA (250)
        RemoveImportTextfile
    End If
    
    Switch.bInline.Visible = False
    Switch.bOutline.Visible = False
    Switch.lblLoad.Caption = "Alpha roster last imported: " & DLookup("LastUpdate", "qLastUpdate")
    'Switch.lblLoad.Visible = False

End Function

Public Function OpenPDF()
'open file dialog to open pdf...
    
    Dim Switch As Form
    Set Switch = Forms!Switchboard
    Switch.lblLoad.Caption = "Opening alpha roster..."
    Switch.Option1.Enabled = False
    Switch.OptionLabel1.Enabled = False


    Dim fd As Object
    'Dim objShell As Object
    
    PDFSelected = False
    'RemoveImportTextfile
    
    Set fd = Application.FileDialog(3)

    'SleepVBA (100)
    
    With fd

        .AllowMultiSelect = False

        ' Set the title of the dialog box.
        .Title = "Open PDF"

        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "PDF", "*.pdf"
        
        If PDFSelected Then Exit Function
        
        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = -1 Then
            
            PDFSelected = True
            'SleepVBA (800)
            
            Switch.lblLoad.Caption = "Converting alpha roster to text..."
            ConvertPDFtoTEXT (.SelectedItems(1))

            'Loop until file is created; probably not the best way to do this
            Do While Dir(GetDBPath & "import.txt") = ""
                SleepVBA 100
                DoEvents
            Loop

            SleepVBA (250)
            Switch.Option1.Enabled = True
            Switch.OptionLabel1.Enabled = True
            OpenPDF = ImportRoster
            Exit Function
        
        Else
            
            Switch.Option1.Enabled = True
            Switch.OptionLabel1.Enabled = True
            Switch.lblLoad.Caption = "Alpha Roster last imported " & DLookup("LastUpdate", "qLastUpdate")

        End If
        
    End With

End Function

Public Function GetDBPath() As String
    
    Dim strFullPath As String
    strFullPath = Mid(DBEngine.Workspaces(0).Databases(0).TableDefs("tCountMain").Connect, 11)
    Dim i As Integer
    
    For i = Len(strFullPath) To 1 Step -1
        
        If Mid(strFullPath, i, 1) = "\" Then
            GetDBPath = Left(strFullPath, i)
            Exit For
        End If
    
    Next

End Function

Public Sub ConvertPDFtoTEXT(ByVal pdf As String)
    
    Dim ConvertCommand As String
    Dim retVal As Double
    ConvertCommand = Chr(34) & GetDBPath & "pdftotext.exe" & Chr(34) & " -simple " & Chr(34) & pdf & Chr(34) & " " & Chr(34) & GetDBPath & "import.txt" & Chr(34)
    retVal = Shell(ConvertCommand, vbMinimizedNoFocus)
    'SleepVBA (1750)

End Sub
