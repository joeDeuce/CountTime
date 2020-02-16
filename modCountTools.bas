Attribute VB_Name = "modCountTools"
Option Compare Database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                            '
'  2019, 2020 Joe Langston, Langston Consulting              '
'                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Global variables updated throughout so
'they can be used anywhere (most notably
'the functions below that are used in queries
Global GlobalOCID As String
Global GlobalCountID As Integer

'Global db object
Global db As Database

'TODO add option to make this configurable
Global Const gCountSlipLines = 67
Global Const gOutCountLines = 28

Option Explicit

Function GetCountSlipLines() As Integer

    GetCountSlipLines = gCountSlipLines

End Function

'Simply returns the global GlobalOCID var value.
'To be used in queries, etc.
Function GetOCID() As String

    GetOCID = GlobalOCID

End Function

'Simply returns the global GlobalCountID var value.
'To be used in queries, etc.
Function GetCountID() As String

    GetCountID = GlobalCountID

End Function

' we need to fill up outcount with spaces if there are less than 30.
' this returns (30 - number of records)
' we will use global var GlobalOCID - just need to ensure
' it is set when applicable
Function GetNumRecordsNeededOutCount()

    Dim Count As Integer
    
    Count = DCount("GDCNum", "tOutCountNames", "OutCountID=" & GlobalOCID)
    
    'If Count is < gOutCountLines then we need (gOutCountLines - Count) empty records to
    'complete the single outcount page
    If Count < gOutCountLines Then
        GetNumRecordsNeededOutCount = gOutCountLines - Count
    Else
        GetNumRecordsNeededOutCount = 0
    End If
    
End Function

'similar to above, this is for count slip
Function GetNumRecordsNeededCountSlip()

    Dim Count As Integer
    Dim Lrs2 As DAO.Recordset
    Dim SQL2 As String
    Set db = CurrentDb()
    SQL2 = "SELECT Count (*) FROM tOutCountNames WHERE (((CountEventID)=" & GetCountID() & ") AND ((DormAtCount) In (SELECT Dorm FROM _tCountSlipPrint WHERE Print=Yes)))"
    Set Lrs2 = db.OpenRecordset(SQL2)
    Count = Lrs2.Fields(0).Value
    
    'If Count is < gCountSlipLines then we need (gCountSlipLines - Count) empty records to
    'complete the single outcount page
    If Count < gCountSlipLines Then
        GetNumRecordsNeededCountSlip = gCountSlipLines - Count
    Else
        GetNumRecordsNeededCountSlip = 0
    End If
    
    db.Close
    Set Lrs2 = Nothing

End Function

'If inmate is already on an outcount, return the
'outcount he is on (returns actual text name)
Function IsOnOutcount(Num As Long, CID As Integer) As String
    
    Dim Lrs2 As DAO.Recordset
    Dim SQL2 As String
    Set db = CurrentDb()
    SQL2 = "SELECT tOutCountLocations.OutCountLocationName " & _
           "FROM tOutCountLocations " & _
           "INNER JOIN (tOutCount " & _
           "INNER JOIN tOutCountNames " & _
           "ON tOutCount.OutCountID = tOutCountNames.OutCountID) " & _
           "ON tOutCountLocations.LocationID = tOutCount.LocationID " & _
           "WHERE (((tOutCount.CountEventID)=" & CID & ") " & _
           "AND ((tOutCountNames.GDCNum)=" & Num & "))"
    Set Lrs2 = db.OpenRecordset(SQL2)
    If Lrs2.EOF And Lrs2.BOF Then
        'not on another outcount
        IsOnOutcount = ""
    Else
        'return loc name
        IsOnOutcount = Lrs2.Fields(0)
    End If
    
    db.Close
    Set Lrs2 = Nothing

End Function

'updates bdoc on worksheet and returns the result as well
Function UpdateBDOC(CountID As String, BDOC As Variant) As String
    
    BDOC.Value = Nz(DCount("GDCNum", "qBDOC"))
    UpdateBDOC = BDOC.Value

End Function

'return OutCountID from CountID and Column
'also create outcount if it doesn't exist, then call self
'to return OutCountID. maybe not needed, but we all need a little recursive love
Function GetOutCountID(CountID, Column) As String
    'THIS IS SOME ORIGINAL CODE FROM BEFORE I KNEW 1/2 OF WHAT I'M DOING
    'PROLLY NEEDS A CLOSE LOOK (i still only know 1/3 what i need to)
    
    Dim Lrs2 As DAO.Recordset
    Dim SQL2 As String
    Set db = CurrentDb()
    SQL2 = "SELECT OutCountID from tOutCount WHERE CountEventID = " & CountID & " AND OutCountColumn = " & Column
    Set Lrs2 = db.OpenRecordset(SQL2)
    If Lrs2.EOF And Lrs2.BOF Then
        'Create New OutCount
        
        'This is fine, but when we close the outcount entry
        'form we need to delete records if nothing entered
        db.Execute "INSERT INTO [tOutCount] ([CountEventID],  [OutCountColumn], [LocationID]) " & _
                   "VALUES (" & CountID & ", " & Column & ",       102)"
        GetOutCountID = GetOutCountID(CountID, Column)
    Else
        GetOutCountID = Lrs2.Fields(0)
    End If
    
    Set Lrs2 = Nothing
    
End Function

'update worksheet values to reflect any changes made
'this is where the majority of things happen
Function UpdateWorksheet(ByRef Worksheet As Form, ByVal CountID As String, Optional ByVal ColumnID As Integer) As Boolean
    On Error GoTo err
    
    Worksheet.lblLoading.Visible = True
    
    GlobalCountID = CountID
    Dim NewBDOC As Integer
    
    ' CountID should equal Worksheet.CountListBox.Value
    ' not entirely convinced this is the best way to do this
    If CountID = -1 Then
        'ADD CODE TO LOAD MOST RECENT RECORD
        '(may no longer be needed)
    Else
        DoCmd.SearchForRecord , "", acFirst, "[CountEventID] = " & CountID
    End If
    
    NewBDOC = UpdateBDOC(CountID, Worksheet.BDOC)
    
    Worksheet.cmdClearTime.Enabled = True
    
    ' create temp table structure that will be used for outcounts and count slips
    Set db = CurrentDb()
    db.Execute "UPDATE [tOutCountUnion] SET CountEventID=" & GlobalCountID
        
    Dim Lrs As Recordset
    Dim SQL As String
    
    Dim LocationTotal As Integer
    
    If IsCountCleared(GlobalCountID) Then
        
        'disable when count cleared
        With Worksheet
            .cmdClearTime.Enabled = False
            .cmbHospitalNames.Enabled = False
            .cmdDeleteHospital.Enabled = False
            .cmdAddHospital.Enabled = False
            .listHospital.Enabled = False
            .Text632.Locked = True
            .Text634.Locked = True
        End With

    Else
    
        'count not cleared, we need to update dorm counts in tCountMain
        Dim DormCountArray(13) As Variant
        Dim AUnitTotal As Integer
        Dim RSATTotal As Integer
        Dim TIC As Integer
        Set db = CurrentDb()
        Dim ID As Integer
        For ID = 1 To 13
            DormCountArray(ID) = DCount("Dorm", "tBaseline", "Dorm = " & ID & " AND InActive = False")
            TIC = TIC + DormCountArray(ID)
            If ID < 8 Or ID > 11 Then
                AUnitTotal = AUnitTotal + DormCountArray(ID)
            Else
                RSATTotal = RSATTotal + DormCountArray(ID)
            End If
        Next
    
        'CALCULATE IN AND OUT TOTALS
        Dim TotIn As Integer
        Dim TotOut As Integer
        TotOut = DCount("GDCNum", "tOutCountNames", "CountEventID = " & GlobalCountID)
        TotIn = TIC - TotOut
    
        'rewrite for M$ bug
        'db.Execute "UPDATE [tCountMain] SET TotalInCount=" & TotIn & ", TotalOutCount=" & TotOut & ", AUnitTotal=" & AUnitTotal & ", RSATTotal=" & RSATTotal & ", TIC=" & TIC & ", UnitCountDorm1=" & DormCountArray(1) & ", UnitCountDorm2=" & DormCountArray(2) & ", UnitCountDorm3=" & DormCountArray(3) & ", UnitCountDorm4=" & DormCountArray(4) & ", UnitCountDorm5 = " & DormCountArray(5) & ", UnitCountDorm6 = " & DormCountArray(6) & ", UnitCountDorm7 = " & DormCountArray(7) & ", UnitCountDorm8 = " & DormCountArray(8) & ", UnitCountDorm9 = " & DormCountArray(9) & ", UnitCountDorm10 = " & DormCountArray(10) & ", UnitCountDorm11 = " & DormCountArray(11) & ", UnitCountDorm12 = " & DormCountArray(12) & ", UnitCountDorm13 = " & DormCountArray(13) & " WHERE CountEventID=" & GlobalCountID
        db.Execute "UPDATE [qCountMain] SET TotalInCount=" & TotIn & ", TotalOutCount=" & TotOut & ", AUnitTotal=" & AUnitTotal & ", RSATTotal=" & RSATTotal & ", TIC=" & TIC & ", UnitCountDorm1=" & DormCountArray(1) & ", UnitCountDorm2=" & DormCountArray(2) & ", UnitCountDorm3=" & DormCountArray(3) & ", UnitCountDorm4=" & DormCountArray(4) & ", UnitCountDorm5 = " & DormCountArray(5) & ", UnitCountDorm6 = " & DormCountArray(6) & ", UnitCountDorm7 = " & DormCountArray(7) & ", UnitCountDorm8 = " & DormCountArray(8) & ", UnitCountDorm9 = " & DormCountArray(9) & ", UnitCountDorm10 = " & DormCountArray(10) & ", UnitCountDorm11 = " & DormCountArray(11) & ", UnitCountDorm12 = " & DormCountArray(12) & ", UnitCountDorm13 = " & DormCountArray(13) & " WHERE CountEventID=" & GlobalCountID


        'enable things needed when count not cleared
        With Worksheet
            .cmdClearTime.Enabled = True
            .cmbHospitalNames.Enabled = True
            .cmdDeleteHospital.Enabled = True
            .cmdAddHospital.Enabled = True
            .listHospital.Enabled = True
            .Text632.Locked = False
            .Text634.Locked = False
        End With
        
    End If
    
    Dim iL As Integer
    Dim LocText As Variant
    Dim LocTotal As Variant
    Dim SkipFor As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim CrossTab As Recordset
    Dim ctCount As Integer
    Dim Column As Integer
    Dim ColumnOut As Integer
    Dim ColumnOutTotal As Integer
    Dim DormID As Integer
    
    Dim CurrentText As Variant
    
    
    ''''''''''''''''''''''' CLEAR THE WORKSHEET ''''''''''''''''''''''''''
    For iL = 1 To 30
        
        Set LocText = Worksheet.Controls("Location" & iL)
        LocText.Caption = "Location " & iL
        '''''LocText.Caption = "|----------->"
        'LocText.FontUnderline = False
        'LocText.FontWeight = vbNormal
        '
        LocText.ForeColor = RGB(0, 175, 225)
        LocText.HyperlinkAddress = "#"
        LocText.ControlTipText = "Add Outcount"
        '
        Set CurrentText = Worksheet.Controls("L" & iL & "Total")
        CurrentText.Value = ""

        For ID = 1 To 13
            
            Set CurrentText = Worksheet.Controls("L" & iL & "D" & ID)
            CurrentText.Value = ""
        
        Next
    
    Next
    
    Set CrossTab = CurrentDb.OpenRecordset("ctJoin_Crosstab")
    ctCount = CrossTab.Fields.Count - 2
    If ctCount = 0 Then
        Worksheet.lblLoading.Visible = False
        GoTo WorksheetContinue
    End If
    
    CrossTab.MoveFirst
    
    While Not CrossTab.EOF
    
        Column = CrossTab.Fields(0)
        ColumnOutTotal = 0
        
        Set LocText = Worksheet.Controls("Location" & Column)
        LocText.Caption = DLookup("OutCountLocationName", "tOutCountLocations Query", "CountEventID = " & GlobalCountID & " AND OutCountColumn = " & Column)
        '
        LocText.ForeColor = RGB(0, 100, 255)
        'LocText.FontUnderline = True
        LocText.ControlTipText = "Edit " & LocText.Caption & " Outcount"
        '

        
        For ID = 1 To ctCount
        
            DormID = CrossTab.Fields(ID + 1).Name
            
            Set CurrentText = Worksheet.Controls("L" & Column & "D" & DormID)
            CurrentText.Value = CrossTab.Fields(ID + 1)
            
        Next
        
        Set CurrentText = Worksheet.Controls("L" & Column & "Total")
        CurrentText.Value = CrossTab.Fields(1)
        
        CrossTab.MoveNext
        
    Wend
    
    CrossTab.Close
        
WorksheetContinue:
        
    Dim GrandSubTotal As Variant
    Dim UC As Variant
    Dim TotalOutText As Variant
    Dim TotalInText As Variant
    Dim TotalTotal As Variant
    Dim TotalOut As Integer
    Dim TotalIn As Integer
    Dim GrandTotal As Integer
    
    'calculate worksheet totals us5 dcoun0?
    For ID = 1 To 13
        SQL = "SELECT Count(*) " & _
              "FROM tOutCountNames " & _
              "WHERE tOutCountNames.CountEventID=" & CountID & " AND tOutCountNames.DormAtCount=" & ID
            
        Set Lrs = db.OpenRecordset(SQL)
           
        Set TotalOutText = Worksheet.Controls("TotalOutD" & ID)
        TotalOutText.Value = Lrs.Fields(0).Value
        If CInt(Lrs.Fields(0).Value) = 0 Then TotalOutText.Value = ""
        
        TotalOut = TotalOut + CInt(Lrs.Fields(0).Value)
        
        Set TotalInText = Worksheet.Controls("TotalInD" & ID)
        Set TotalTotal = Worksheet.Controls("UCD" & ID)
        TotalInText.Value = CInt(TotalTotal.Value) - Lrs.Fields(0).Value
        
        TotalIn = TotalIn + CInt(TotalInText.Value)
    Next

    If TotalOut > 0 Then Worksheet.TotalOutTotal.Value = TotalOut Else Worksheet.TotalOutTotal.Value = ""
    Worksheet.TotalInTotal.Value = TotalIn
    
    Worksheet.cmdDelete.Enabled = True
    Worksheet.cmdDuplicate.Enabled = True
    Worksheet.cmdReloadWorksheet.Enabled = True
    Worksheet.cmdPrintPackage.Enabled = True
    Worksheet.CountListBox.Requery
    Worksheet.listHospital.Visible = True
    Worksheet.listHospital.Requery
    
Bye:
    
    If Worksheet.Dirty Then Worksheet.Dirty = False
    Worksheet.listHospital.Requery
    
    db.Close
    Set db = Nothing
    Set Lrs = Nothing
    Set LocText = Nothing
    Set LocTotal = Nothing
    Set CurrentText = Nothing
    Set TotalOutText = Nothing
    Set TotalTotal = Nothing
    Set TotalInText = Nothing
    
    UpdateWorksheet = True
    Worksheet.lblLoading.Visible = False
    
    Exit Function

err:
    MsgBox Error$
    Resume Bye

End Function

Function ReportWorksheet(ByRef Worksheet As Report, ByVal CountID As Integer, Optional ByVal Blank As Boolean) As Boolean
    
    Dim NewBDOC As Integer
    
    'Choose count based on drop down selection
    DoCmd.SearchForRecord , "", acFirst, "[CountEventID] = " & CountID

    Dim Lrs As DAO.Recordset
    Dim SQL As String
    
    Set db = CurrentDb()
    
    Dim LocationTotal As Integer
    Dim iL As Integer
    Dim ID As Integer
    Dim LocText As Variant
    Dim LocTotal As Variant
    
    Dim CrossTab As Recordset
    Dim ctCount As Integer
    Dim Column As Integer
    Dim ColumnOut As Integer
    Dim ColumnOutTotal As Integer
    Dim DormID As Integer
    
    Dim CurrentText As Variant
    
    Set CrossTab = CurrentDb.OpenRecordset("ctJoin_Crosstab")
    ctCount = CrossTab.Fields.Count - 2
    If ctCount = 0 Then
        GoTo ReportContinue
    End If
    
    CrossTab.MoveFirst
    
    While Not CrossTab.EOF
    
        Column = CrossTab.Fields(0)
        ColumnOutTotal = 0
        
        Set LocText = Worksheet.Controls("Location" & Column)
        LocText.Value = DLookup("OutCountLocationName", "tOutCountLocations Query", "CountEventID = " & GlobalCountID & " AND OutCountColumn = " & Column)
        
        For ID = 1 To ctCount
        
            DormID = CrossTab.Fields(ID + 1).Name
            
            Set CurrentText = Worksheet.Controls("L" & Column & "D" & DormID)
            CurrentText.Value = CrossTab.Fields(ID + 1)
            
        Next
        
        Set CurrentText = Worksheet.Controls("L" & Column & "Total")
        CurrentText.Value = CrossTab.Fields(1)
        
        CrossTab.MoveNext
        
    Wend
    
    CrossTab.Close
    
ReportContinue:

    Dim GrandSubTotal As Variant
    Dim UC As Variant
    Dim TotalOutText As Variant
    Dim TotalInText As Variant
    Dim TotalTotal As Variant
    Dim TotalOut As Integer
    Dim TotalIn As Integer
    Dim GrandTotal As Integer
    
    For ID = 1 To 13
        SQL = "SELECT Count(*) " & _
              "FROM tOutCountNames " & _
              "WHERE tOutCountNames.CountEventID=" & CountID & " AND tOutCountNames.DormAtCount=" & ID
            
        Set db = CurrentDb()
        Set Lrs = db.OpenRecordset(SQL)
              
        Set GrandSubTotal = Worksheet.Controls("GT" & ID)
        Set UC = Worksheet.Controls("UCD" & ID)
        Set TotalOutText = Worksheet.Controls("TotalOutD" & ID)
        Set TotalInText = Worksheet.Controls("TotalInD" & ID)
        Set TotalTotal = Forms!fWorksheet.Controls("UCD" & ID)
        
        'number out
        TotalOutText.Value = Lrs.Fields(0).Value
        
        'total out
        TotalOut = TotalOut + CInt(Lrs.Fields(0).Value)
        
        'number in
        TotalInText.Value = CInt(TotalTotal) - Lrs.Fields(0).Value
        
        'grand dorm / unit count
        GrandSubTotal.Value = TotalOutText.Value + TotalInText.Value
        UC.Value = GrandSubTotal.Value
        
        'total in
        TotalIn = TotalIn + Int(TotalInText.Value)
        
        'grand totals (include hospitals)
        GrandTotal = GrandTotal + UC.Value
        
        'Change 0s to "" in total out
        If Lrs.Fields(0).Value = 0 Then TotalOutText.Value = ""
        
        'Error checking, although I don't see the possibililty of this happening
        If TotalTotal <> UC.Value Then
            MsgBox "Calculation Error!", vbCritical, "Error!"
            ReportWorksheet = False
        End If
    Next
    
    Worksheet.UCTIC.Value = GrandTotal + GetHospitalCount(CountID)
    Worksheet.GrandTotal.Value = Worksheet.UCTIC.Value
        
    Set db = Nothing
    Set Lrs = Nothing
    Set LocText = Nothing
    Set LocTotal = Nothing
    Set CurrentText = Nothing
    Set TotalOutText = Nothing
    Set TotalTotal = Nothing
    Set TotalInText = Nothing

    If TotalOut > 0 Then Worksheet.TotalOutTotal.Value = TotalOut Else Worksheet.TotalOutTotal.Value = ""
    Worksheet.TotalInTotal.Value = TotalIn + GetHospitalCount(CountID)

Bye:
    If Blank Then
        Dim ctl As Control
        For Each ctl In Worksheet.Controls
            If ctl.Tag = "ClearOnBlank" Then
                ctl.ForeColor = vbWhite
            End If
        Next
    End If
    
    ReportWorksheet = True
    Exit Function

err:
    MsgBox Error$
    Resume Bye

End Function

'Return the number of inmates out to hospital
Function GetHospitalCount(CID As Integer) As Integer
    
    GetHospitalCount = Nz(DCount("*", "tHospitalMembers", "CountEventID=" & CStr(CID)), 0)
    
End Function

'returns true if the count cleared time has been entered
Function IsCountCleared(CID As Integer) As Boolean

    Dim Lrs2 As DAO.Recordset
    Dim SQL2 As String
    
    Set db = CurrentDb()
    SQL2 = "SELECT CountCleared FROM tCountMain WHERE (CountEventID=" & CStr(CID) & ")"
    Set Lrs2 = db.OpenRecordset(SQL2)
    
    If Not Lrs2.Fields(0).Value Then
        IsCountCleared = True
    Else
        IsCountCleared = False
    End If
    
    Set Lrs2 = Nothing

End Function

Public Sub HideNavPane(bVisible As Boolean)
    On Error GoTo Error_Handler
 
    If SysCmd(acSysCmdRuntime) = False Then
        If bVisible = True Then
            DoCmd.SelectObject acModule, , True
        Else
            DoCmd.SelectObject acModule, , True
            DoCmd.RunCommand acCmdWindowHide
        End If
    End If
 
Error_Handler_Exit:
    Exit Sub
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: HideNavPane" & vbCrLf & _
           "Error Description: " & err.Description _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
    
End Sub

Function GetFileName(FullPath As String) As String

    Dim splitList As Variant
    splitList = VBA.Split(FullPath, "\")
    GetFileName = splitList(UBound(splitList, 1))

End Function

Function DuplicateCompleteCount(OldID As Integer, ByRef Mee As Variant) As Boolean
    
    Dim DormCountArray(13) As Variant
    Dim AUnitTotal As Integer
    Dim RSATTotal As Integer
    Dim TIC As Integer
    Dim NewCountID As String
    Dim NumAffected As Integer

    'code that creates new, blank count.
    'TODO needs to be consolidated with similar code in Form_fCreateNew module
    Dim ID As Integer
    Dim db2 As Database
    Dim Lrs2 As DAO.Recordset
    Dim SQL2 As String
    Set db2 = CurrentDb()
    
    For ID = 1 To 13

        Dim Lrs As DAO.Recordset
        Dim SQL As String
        Set db = CurrentDb()

        SQL = "SELECT Count(*) AS DormCount FROM tBaseline WHERE (((tBaseline.Dorm)=" & ID & ") AND ((tBaseline.InActive)=False));"
    
        Set Lrs = db.OpenRecordset(SQL)
        DormCountArray(ID) = Lrs.Fields(0).Value
        TIC = TIC + DormCountArray(ID)
        If ID < 8 Or ID > 11 Then
            AUnitTotal = AUnitTotal + DormCountArray(ID)
        Else
            RSATTotal = RSATTotal + DormCountArray(ID)
        End If
        
        db.Close
        Set Lrs = Nothing
    
    Next
    
    'create blank count with old count's basic data
    Set db = CurrentDb()
    db.Execute "INSERT INTO [tCountMain] (CountDate, CountTime, AUnitTotal, RSATTotal, TIC, UnitCountDorm1, UnitCountDorm2, UnitCountDorm3, UnitCountDorm4, UnitCountDorm5, UnitCountDorm6, UnitCountDorm7, UnitCountDorm8, UnitCountDorm9, UnitCountDorm10, UnitCountDorm11, UnitCountDorm12, UnitCountDorm13) VALUES " & _
               "(#" & Mee.NewDate & "#, #" & Mee.NewTime & "#, " & AUnitTotal & ", " & RSATTotal & ", " & TIC & ", " & DormCountArray(1) & ", " & DormCountArray(2) & ", " & DormCountArray(3) & ", " & DormCountArray(4) & ", " & DormCountArray(5) & ", " & DormCountArray(6) & ", " & DormCountArray(7) & ", " & DormCountArray(8) & ", " & DormCountArray(9) & ", " & DormCountArray(10) & ", " & DormCountArray(11) & ", " & DormCountArray(12) & ", " & DormCountArray(13) & ")"
    
    'retrieve new CountEventID
    SQL = "SELECT TOP 1 CountEventID FROM tCountMain ORDER BY CountEventID DESC"
    Set Lrs = db.OpenRecordset(SQL)
    NewCountID = Lrs.Fields(0).Value
   
    db.Execute "DELETE * FROM _tOutCountNames"
    
    db.Execute "INSERT INTO [tOutCount] (CountEventID, OutCountColumn, LocationID) SELECT '" & NewCountID & "' AS CountEventID, OutCountColumn, LocationID FROM tOutCount WHERE CountEventID = " & OldID
    NumAffected = db.RecordsAffected
    
    'now we need to pull the newest OutCountIDs from tOutCount
    'so that we can link the outcounts with the locations
    'we should be able to match them up with columns
    SQL = "SELECT TOP " & NumAffected & " OutCountID FROM tOutCount ORDER BY OutCountID DESC"
    Set Lrs = db.OpenRecordset(SQL)
    
    SQL2 = "SELECT OutCountID FROM tOutCount WHERE CountEventID=" & OldID & " ORDER BY OutCountID DESC"
    Set Lrs2 = db2.OpenRecordset(SQL2)
    
    Dim ColumnCount As Integer
    
    For ColumnCount = 1 To NumAffected
        db.Execute "INSERT INTO [tOutCountNames] (CountEventID, OutCountID, GDCNum, DormAtCount) SELECT " & NewCountID & " AS CountEventID, " & Lrs.Fields(0).Value & " AS OutCountID, GDCNum, DormAtCount FROM tOutCountNames WHERE OutCountID = " & Lrs2.Fields(0).Value
        Lrs.MoveNext
        Lrs2.MoveNext
    Next
    
    'copy hospitals
    db.Execute "INSERT INTO [tHospitalMembers] (CountEventID, HospitalID, GDCNum) SELECT " & NewCountID & " AS CountEventID, HospitalID, GDCNum FROM tHospitalMembers WHERE CountEventID = " & OldID
        
    db.Execute "UPDATE tBaseline INNER JOIN tOutCountNames ON tBaseline.GDCNum = tOutCountNames.GDCNum SET tOutCountNames.DormAtCount = tBaseline.Dorm WHERE (tOutCountNames.CountEventID = " & NewCountID & ")"
        
    db.Execute "5qSanitizeOpenCounts"

    DoCmd.Close acForm, "fWorksheet", acSaveYes
    DoCmd.OpenForm "fWorksheet", acNormal, "", "", , acNormal

End Function

Function GetCurrentDorm(InmateNumber As Long) As Integer

    GetCurrentDorm = DLookup("[Dorm]", "[tBaseline]", "GDCNum = " & InmateNumber)

End Function

'clean any counts older than 3 years
'TODO add option to adjust age of counts that get removed
Function CleanOldCounts() As Boolean
    
    Dim RowsAffected As Integer
    Set db = CurrentDb()
    RowsAffected = DCount("CountEventID", "tCountMain", "DATEADD('yyyy', 3, CountDate) < date()")
    
    If RowsAffected > 0 Then
        If MsgBox("You are about to DELETE " & RowsAffected & " counts that are over 3 years old. Are you SURE you want to do this?", vbYesNo, "Delete Counts") = vbYes Then
            db.Execute "DELETE * FROM tCountMain WHERE DATEADD('yyyy', 3, CountDate) < date()"
            MsgBox "Deleted " & db.RecordsAffected & " counts.", vbOKOnly, "Success"
        End If
    ElseIf RowsAffected = 0 Then
        MsgBox "There is nothing to delete!", vbOKOnly, "Delete Counts"
    End If
    
End Function

Function IsCountBlank(CIDD As Integer) As Boolean

    If DCount("GDCNum", "tOutCountNames", "CountEventID = " & CIDD) > 0 Then
        IsCountBlank = False
    Else
        IsCountBlank = True
    End If
    
End Function

Function GetOption(Name As String) As Integer

    GetOption = DLookup("OptionValue", "_tOptions", "OptionName='" & Name & "'")

End Function

Sub SetOption(Name As String, Opt As Integer)
            
    Set db = CurrentDb()
    db.Execute "UPDATE _tOptions SET OptionValue=" & Opt & " WHERE OptionName='" & Name & "'"

End Sub

Sub RemoveImportTextfile()

    Dim del As String
    del = GetDBPath & "import.txt"
    If FileExists(del) Then
        ' First remove readonly attribute, if set
        SetAttr del, vbNormal
        ' Then delete the file
        On Error Resume Next
        Kill del
        'On Error GoTo 0
    End If

End Sub


