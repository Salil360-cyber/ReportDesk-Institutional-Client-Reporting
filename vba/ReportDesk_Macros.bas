Attribute VB_Name = "ReportDeskMacros"
Option Explicit

'==============================================================================
' ReportDesk - VBA macro
'
'   GenerateClientReports   For every client in Data_Summary:
'                             - READY clients: puts the client ID in Settings!B2,
'                               recalculates, and saves the Report sheet as a PDF
'                             - BLOCKED clients: skipped (data issue must be fixed)
'                           Every client gets a line in Report_Log.
'
' PDFs are saved in a "Client_Reports" folder next to this workbook.
'==============================================================================

Private Function ColOf(ws As Worksheet, headerText As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then
        Err.Raise vbObjectError + 513, "ReportDesk", "Column not found on " & ws.Name & ": " & headerText
    End If
    ColOf = f.Column
End Function

Private Sub WriteLog(wsLog As Worksheet, clientId As String, clientName As String, _
                     dataStatus As String, result As String, fileName As String)
    Dim r As Long
    r = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(r, 1).Value = Now
    wsLog.Cells(r, 1).NumberFormat = "dd-mmm-yyyy hh:mm"
    wsLog.Cells(r, 2).Value = clientId
    wsLog.Cells(r, 3).Value = clientName
    wsLog.Cells(r, 4).Value = dataStatus
    wsLog.Cells(r, 5).Value = result
    wsLog.Cells(r, 6).Value = fileName
    wsLog.Range(wsLog.Cells(r, 1), wsLog.Cells(r, 6)).Font.Name = "Arial"
End Sub

Public Sub GenerateClientReports()
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first, so the reports folder can be created next to it.", _
               vbExclamation, "ReportDesk"
        Exit Sub
    End If

    Dim wsData As Worksheet, wsRep As Worksheet, wsSet As Worksheet, wsLog As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data_Summary")
    Set wsRep = ThisWorkbook.Worksheets("Report")
    Set wsSet = ThisWorkbook.Worksheets("Settings")
    Set wsLog = ThisWorkbook.Worksheets("Report_Log")

    Dim cId As Long, cName As Long, cStatus As Long
    cId = ColOf(wsData, "client_id")
    cName = ColOf(wsData, "client_name")
    cStatus = ColOf(wsData, "report_status")

    Dim folder As String, monthLabel As String
    folder = ThisWorkbook.Path & Application.PathSeparator & "Client_Reports"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    monthLabel = Format(wsSet.Range("B1").Value, "yyyy-mm")

    Dim originalClient As Variant
    originalClient = wsSet.Range("B2").Value

    Dim lastRow As Long, r As Long, made As Long, skipped As Long
    Dim clientId As String, clientName As String, status As String, fileName As String
    lastRow = wsData.Cells(wsData.Rows.Count, cId).End(xlUp).Row

    Application.ScreenUpdating = False
    For r = 2 To lastRow
        clientId = CStr(wsData.Cells(r, cId).Value)
        clientName = CStr(wsData.Cells(r, cName).Value)
        status = CStr(wsData.Cells(r, cStatus).Value)

        If Left(status, 5) = "READY" Then
            wsSet.Range("B2").Value = clientId
            Application.Calculate
            fileName = folder & Application.PathSeparator & clientId & "_Portfolio_Report_" & monthLabel & ".pdf"
            wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, _
                                      IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            If status = "READY" Then
                WriteLog wsLog, clientId, clientName, status, "PDF created", fileName
            Else
                WriteLog wsLog, clientId, clientName, status, "PDF created - confirm warnings before sending", fileName
            End If
            made = made + 1
        Else
            WriteLog wsLog, clientId, clientName, status, "Skipped - fix BLOCKER in Data_DQ first", ""
            skipped = skipped + 1
        End If
    Next r

    wsSet.Range("B2").Value = originalClient
    Application.Calculate
    Application.ScreenUpdating = True
    wsLog.Activate

    MsgBox made & " report(s) created in the Client_Reports folder." & vbCrLf & _
           skipped & " client(s) skipped because of BLOCKER data issues.", vbInformation, "ReportDesk"
End Sub
