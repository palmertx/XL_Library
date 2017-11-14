    Option Explicit

Dim sSP As String
Dim iDBOrg As Integer
Dim iDBContactID As Long
Dim iTemp As Integer
Dim vTemp As Variant
Dim iYear As Integer
Dim dtLastMonth As Date
Dim vSet As Worksheet
Dim sCommand, sConnStr, sRptHeader As String
Dim sDBOrgName As String
Dim sDBName As String
Dim iRow_DB As Integer
Dim iRows_DB As Integer
Dim sPathRoot As String
Dim sPathDB As String
Dim sFilename As String
Dim sFullFileName As String
Dim sLastFullFileName As String
Dim sStatusBar As String
Dim bInit As Boolean
Dim bInitOL As Boolean
Dim vWshtDBs As Worksheet
Dim iColCaseCnt As Integer
Dim iColContractCnt As Integer
Dim iColFileLink As Integer
Dim iColContractCode As Integer
Dim iColStatusCode As Integer
Dim sTemp As String
Dim vOL As Object
Dim olNS As Object
Dim sMe As String
Dim vDB As Worksheet
Dim iCaseID As Long
Dim vWshtClieynts As Worksheet
Dim sReportID As String
Public sEmailTemplate As String
Public bNoDrafts As Boolean
Dim wbkDB As Workbook
Dim shtDB As Worksheet

Sub Init()
  If bInit Then Exit Sub
  Set vSet = ThisWorkbook.Worksheets("Settings")
  sTemp = DNStr("Workbook.Name")
  Set wbkDB = Workbooks(sTemp)
  Set shtDB = wbkDB.Worksheets(DNStr("Worksheet.Name"))
  With Application
     .StandardFont = "Arial"
     .StandardFontSize = "10"
     .SheetsInNewWorkbook = 1
  End With
  'bInit = True
End Sub

Sub Auto_Open()
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets("Settings").Visible = True
    ThisWorkbook.Worksheets("Settings").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Range("Sheet").Select
    ActiveWindow.Zoom = True  'Adjust for Screen Resolution
    Range("A1").Select        'Hide Cursor
    Application.ScreenUpdating = True
 End Sub


Sub InitOL()
   bInit = False
   Call Init
   If bInitOL Then Exit Sub
   Set vOL = CreateObject("Outlook.Application")
'    If Empty = vOL Then
'       MsgBox "Open MS Outlook first"
'       End
'    End If
    Set olNS = vOL.GetNamespace("MAPI")
    sMe = olNS.currentuser
   'If MsgBox(sM, vbYesNo) = vbNo Then Stop
    If Len(Application.UserName) <= 4 And Len(sMe) >= 10 Then Application.UserName = sMe
    bInitOL = True
End Sub

Sub SendEmails()
   Dim sEmail As String
   Dim sFirstName As String
   Dim sLastname As String
   Dim sM As String
   Dim vMessage As Object
   Dim iEmailCnt As Integer
   
   Call InitOL
   sEmailTemplate = ""
 
   iRows_DB = vSet.Range("iRows_DB").Item(1, 1).Value
   'Rows_DB = 94
   For iRow_DB = 2 To 1000
      vSet.Range("iRow.DB").Item(1, 1).Value = iRow_DB
      Application.Calculate
      If vSet.Range("Max.Text").Value < 4 Then Exit For
      Call SendAnEmail
      iEmailCnt = iEmailCnt + 1
   Next iRow_DB
   Application.StatusBar = False
   sTemp = iEmailCnt & " of the " & iRows_DB & " DBs had case loads found and emails were generated.  "
   sTemp = sTemp & " If this number seems low, make sure that ALL the Excel files where generated before running the email generator."
   sTemp = sTemp & "  Review Column '" & Chr(64 + iColCaseCnt) & "' in the '" '& vWshtDBs.Name & "' sheet/tab."
   Beep
   MsgBox sTemp, vbOKOnly, "Results"

End Sub



Sub SendAnEmail()
   Dim sEmail As String
   Dim sFile As String
   Dim bListenHTM As Boolean
   Dim sFirstName As String
   Dim sLastname As String
   Dim sM As String
   Dim sSubject As String
   Dim vMessage As Object
   Application.Calculate
   Application.StatusBar = "Initializing Outlook Connection"
   Call InitOL
   If Len(sEmailTemplate) = 0 Then
      sFile = DNStr("Template.Path") & "\" & vSet.Range("File.HTML").Text
      Application.StatusBar = "Opening: " & sFile
      Close
      Open sFile For Input As #1
      bListenHTM = False
      While Not EOF(1)
         Line Input #1, sTemp
         If InStr(LCase(sTemp), "<body") >= 1 Then bListenHTM = True
         If InStr(LCase(sTemp), "</body") >= 1 Then bListenHTM = False
         If bListenHTM Then
            sEmailTemplate = sEmailTemplate & vbCrLf & sTemp
         End If
      Wend
   End If
   
       vTemp = Application.Match("FirstName", vSet.Range("Fields"), 0)
       If IsError(vTemp) Then vTemp = Application.Match("First Name", vSet.Range("Fields"), 0)
       If IsError(vTemp) Then vTemp = Application.Match("First", vSet.Range("Fields"), 0)
       iTemp = vSet.Range("Fields").Row + vTemp - 1
       sTemp = vSet.Cells(iTemp, 2).Text
       vTemp = Application.Match("LastName", vSet.Range("Fields"), 0)
       If IsError(vTemp) Then vTemp = Application.Match("Last Name", vSet.Range("Fields"), 0)
       If IsError(vTemp) Then vTemp = Application.Match("Last", vSet.Range("Fields"), 0)
       iTemp = vSet.Range("Fields").Row + vTemp - 1
       sTemp = sTemp & " " & vSet.Cells(iTemp, 2).Text
       
       iTemp = Application.Match("email", vSet.Range("Fields"), 0)
       iTemp = vSet.Range("Fields").Row + iTemp - 1
       sEmail = sTemp & " [" & vSet.Cells(iTemp, 2).Text & "]"
       sStatusBar = Format((iRow_DB - 2) / (iRows_DB - 1), "0%") & " Email for " & sDBOrgName
       Application.StatusBar = sStatusBar
       If Len(vSet.Cells(iTemp, 2).Text) >= 4 And InStr(sEmail, "@") >= 2 Then
       '--------- Create Mail Item --------
       sM = FormFill()
       Set vMessage = vOL.CreateItem(0)  'olMailItem
       If Len(sEmail) = 0 Then sEmail = sMe
       With vMessage
          With .Recipients.Add(sEmail)
               .Type = 1 ' '=olTo
          End With
          sTemp = Trim(vSet.Range("Email.CC").Text)
          If Len(sTemp) >= 3 Then
             With .Recipients.Add(sTemp)
                 .Type = 2 '=olCC
             End With
          End If
          sTemp = Trim(vSet.Range("Email.BCC").Text)
          If Len(sTemp) >= 3 Then
             With .Recipients.Add(sTemp)
                 .Type = 3 '=olBCC
             End With
          End If
          sSubject = Trim(vSet.Range("Email.Subject").Text)
          .Subject = sSubject
          '.Body = sM
          .BodyFormat = 2 'Outlook.OlBodyFormat.olFormatHTML ' 2 'olFormatHTML
          .HTMLBody = sM

          'If DNBool("email.WithXLS") Then
              'With .Attachments.Add(sFullFileName, 1) '4=olByReference 5=olEmbeddedItem 1=olByValue
                '.DisplayName = .Subject
              'End With
          'End If
          If bNoDrafts Then
            .Send
          Else
            .Save
          End If
       End With
     End If 'Valid Email
    Application.StatusBar = False

End Sub


Sub MakeNewDataSheet(sSht As String)
Dim sShtname As String
Dim sShtOld As String
Exit Sub
' Macro1 Macro
' Macro recorded 11/8/2005 by bxak
'
'MakeNewDataSheet ("DBContacts")
    sShtOld = sSht & "_Old"
    
    'Sheets(sShtOld).Delete
   
    Sheets(sSht).Name = sShtOld
    Sheets.Add
    sShtname = ActiveSheet.Name
    Sheets(sShtOld).Select
    Cells.Select
    Selection.Copy
    Sheets(sShtname).Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets(sShtOld).Select
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets(sShtname).Select
    Rows("1:1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(sShtname).Select
    Sheets(sShtname).Name = sSht
    Range("A1").Select
    RunQuery ("sp_Select_Report_862_Recipients")

End Sub


  Function FormFill() As String
    Dim sLetter As String
    Dim sFind, sReplace As String
    Dim iFields As Integer, iField As Integer
    Dim sField As String
    Dim iReplaces As Integer
    Dim sFields As String
    
    sLetter = sEmailTemplate
    iFields = vSet.Range("Fields").Rows.Count
    For iField = 1 To iFields
      sField = vSet.Range("Fields").Item(iField).Text
      If Len(sField) >= 2 Then
        sFind = "{" & sField & "}"
        'If sField = "and" Then Stop
        If InStr(sLetter, sFind) >= 1 Then  'Found Database Item in text?
           sReplace = vSet.Range("Data").Item(iField).Text
           If Len(sReplace) >= 1 Then
              sReplace = Replace(sReplace, "<br> <br>", "<br>")
           Else
              sReplace = " " '&nbsp; & sFind
           End If
          iReplaces = iReplaces + 1
          sLetter = Trim(Replace(sLetter, sFind, sReplace))
        End If
      sFields = sFields & ", " & sField
      End If
    Next iField
    FormFill = sLetter
    
  End Function

