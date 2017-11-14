Option Explicit

Function TrackID(ByVal iNumber As Integer) As String
   If iNumber >= 10 Then
      TrackID = Chr(55 + iNumber)
   Else
      TrackID = CStr(iNumber)
   End If
End Function


Function RunQuery(ByVal sSQL As String, Optional ByVal vWbk As Workbook, Optional ByVal vWkst As Worksheet, Optional ByVal sCell As String, Optional ByVal sConnStr As String) As String
   'Const sCell = "A1"
   'Const sConnStr = "ODBC;DRIVER=SQL Server;SERVER=MAS3;APP=Microsoft Office 2003;Trusted_Connection=Yes"
   'Dim vWkst As Worksheet
   'Dim vWbk As Workbook
   'Set vWkst = ActiveSheet
   'Set vWbk = ThisWorkbook
   If Len(sCell) = 0 Then sCell = "A1"
   If Len(sConnStr) = 0 Then sConnStr = "ODBC;DRIVER=SQL Server;SERVER=MAS3;DATABASE=MAS;APP=Microsoft Office 2003;Trusted_Connection=Yes"
   If Not IsMissing(vWbk) Then
      'ActiveWorkbook.Activate
   Else
      vWbk.Activate
   End If
   If Not IsMissing(vWkst) Then
      'vWkst = vWbk.ActiveSheet
   Else
      vWkst.Select
      vWkst.Activate
   End If
   '---- Clear Previous Queries -----
   Dim sTemp As String, iQuery As Integer
   For iQuery = ActiveSheet.QueryTables.Count To 1 Step -1
      sTemp = ActiveSheet.QueryTables(iQuery).Name
      ActiveSheet.QueryTables(iQuery).Delete
   Next iQuery
   Dim sName As String
   sName = "Query from MAS"
   If Left(sSQL, 3) = "sp_" Then sName = sSQL
   If InStr(sName, " ") Then sName = Trim(Left(sName, InStr(sName, " ")))

    'vWbk.Activate
    'vWkst.Activate
    Range(sCell).Select
    'sConnStr = "ODBC;DRIVER=SQL Server;SERVER=MAS3;APP=Microsoft Office XP;DATABASE=MAS;Trusted_Connection=Yes" _
   '
   '     With ActiveSheet.QueryTables.Add(Connection:=sConnStr, Destination:=Range(sCell))
   '     .CommandText = Array(sSQL) '"sp_Select_Report_862_Recipients")
   '     .Name = sName
   '     .FieldNames = True
   '     .RowNumbers = False
   '     .BackgroundQuery = True
   '     .RefreshStyle = xlInsertDeleteCells
   '     .SavePassword = False
   '     .SaveData = False
   '     .AdjustColumnWidth = True
   '     .RefreshPeriod = 0
   '     .PreserveColumnInfo = False
   '     .Refresh BackgroundQuery:=False
   ' End With
    
    With ActiveSheet.QueryTables.Add(Connection:=sConnStr, Destination:=Range(sCell))
        
        .CommandText = Array(sSQL)
        .Name = sName
        .FieldNames = True
        .RowNumbers = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = False
        .Refresh BackgroundQuery:=False
    End With
    RunQuery = "Successful"
End Function

Function FileExists(ByVal sFullname As String) As Boolean
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")
   FileExists = fso.FileExists(sFullname)
End Function

Function FolderExists(ByVal sFullname As String, ByVal bMkdir As Boolean) As Boolean
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")
   If InStr(sFullname, ".") > InStr(sFullname, "\") Then
      sFullname = Left(sFullname, InStrRev(sFullname, "\"))
   End If
   FolderExists = fso.FolderExists(sFullname)
   If bMkdir And FolderExists = False Then
      fso.CreateFolder sFullname
      FolderExists = True
   End If
   
End Function

Function FileOpened(ByVal sFullname As String) As Boolean
   FileOpened = False
   Dim wbk As Workbook
   For Each wbk In Workbooks
      If InStr(UCase(sFullname), UCase(wbk.Name)) >= 1 Then
         FileOpened = True
      End If
   Next wbk
End Function

Function FileName(sFullname As Variant, iType As Variant) As Variant

'Where iType can choose:
'    1  = FileName in "FILENAME.EXT" format
'    2  = PathName = X:\PATH\PATH" format
'    3  = Drive Letter in "X:" format
'    4  = File Extension in ".EXT" format
'    5  = File Root in "FILENAME" format
'    6  = Filename in Full in "X:\PATH\PATH\FILENAME.EXT" format
'    7 = Test for file's Existence on disk
'    8 = Test to See if file/workbook is Open
'   10 = File Date in mm/dd/yy format
'   11 = File Date in English  "dddd, mmmm d, yyyy" format
'   12 = File Time in "hh:mm" format
   
   Dim sPart(7) As String, sRemains As String, iFound As Integer
   Dim bError As Boolean, dtFile As Date, wb As Workbook
      
   '-------Find Drive Letter------
   sFullname = Trim(sFullname)
   If Mid(sFullname, 2, 1) = ":" Then
      sPart(1) = Left(sFullname, 1)
      sRemains = Mid(sFullname, 3, 120)
   ElseIf Left(sFullname, 2) = "\\" Then
      sPart(1) = ""
      sRemains = sFullname
   Else
      sPart(1) = Left(CurDir(), 1)
      sRemains = sFullname
   End If
   '-------Find File Extension ------
   iFound = InStr(Right(sRemains, 4), ".")
   If iFound > 0 Then
      sPart(4) = Mid(Right(sRemains, 4), iFound, 4)
      sRemains = Left(sRemains, Len(sRemains) - Len(sPart(4)))
   Else
      sPart(4) = ".   "
   End If
   '------- Find Path & Root File Name -------
   iFound = 0
   While InStr(iFound + 1, sRemains, "\") > 0
      iFound = InStr(iFound + 1, sRemains, "\")
   Wend
   If iFound = 0 Then
      sPart(2) = Mid(CurDir(), 3, 120)
      sPart(3) = sRemains
   Else
      sPart(2) = Left(sRemains, iFound - 1)
      sPart(3) = Mid(sRemains, iFound + 1, 32)
   End If
   If Left(sPart(2), 1) <> "\" Then sPart(2) = "\" & sPart(2)
   sFullname = sPart(2) & "\" & sPart(3) & sPart(4)
   If Len(sPart(1)) >= 1 Then sFullname = sPart(1) & ":" & sFullname
   '--------- File Data ------------------
   Select Case iType
   Case 1  'FileName in "FILENAME.EXT" format
       FileName = sPart(3) & sPart(4)
   Case 2  'PathName = X:\PATH\PATH" format
       FileName = sPart(2)
       If Len(sPart(1)) >= 1 Then FileName = sPart(1) & ":" & FileName
   Case 3  'Drive Letter in "X:" format
       FileName = sPart(1)
   Case 4  'File Extension in ".EXT" format
       FileName = sPart(4)
   Case 5  'File Root in "FILENAME" format
       FileName = sPart(3)
   Case 6  'Filename in Full in "X:\PATH\PATH\FILENAME.EXT" format
       FileName = sFullname
   Case 7 'Test for file's Existence on disk
       FileName = True
       On Error GoTo BadPath
       If Len(Trim(sPart(3))) = 0 Then bError = True
       iFound = Len(Dir(sFullname))
       On Error GoTo 0
       If bError Or iFound = 0 Then FileName = False
   Case 8 'Test to See if file/workbook is Open
       FileName = False
       For Each wb In Workbooks
          If UCase(wb.Name) = UCase(sPart(3) & sPart(4)) Then
             FileName = True
             Exit For
          End If
       Next wb
   Case 10 'File Date in "mm/dd/yy" format
       dtFile = FileDateTime(sFullname)
       sPart(6) = Format(dtFile, "mm/dd/yy")
       FileName = sPart(6)
   Case 11 'File Date in English  "dddd, mmmm d, yyyy" format
       dtFile = FileDateTime(sFullname)
       sPart(6) = Format(dtFile, "dddd, mmmm d, yyyy")
       FileName = sPart(6)
   Case 12 'File Time in "hh:mm" format
       dtFile = FileDateTime(sFullname)
       sPart(7) = Format(dtFile, "hh:mm")
       FileName = sPart(7)
   Case Else
       FileName = "Invalid Type Request"
   End Select
   Exit Function

'------- Test for Existence ----
FileExist:
Return

BadPath:
   bError = True
   Resume Next
Return
   
End Function

'----------- Fetch a Defined Name Value that is Bookean -------------------
Function DNBool(s As String) As Boolean
   If InStr(vMainWbk.Names(s), "TRUE") >= 2 Then DNBool = True Else DNBool = False
End Function
'----------- Fetch a Defined Name Value that is String -------------------

Function DNStr(s As String) As String
   Dim ss As String
   Dim iPos As Integer
   ss = ThisWorkbook.Names(s)
   If Left(ss, 2) <> "=" & Chr(34) Then
      DNStr = Mid(ss, 2, Len(ss) - 1)
   Else
      DNStr = Mid(ss, 3, Len(ss) - 3)
   End If
   iPos = InStr(DNStr, "!$")
   If iPos >= 2 Then
      ss = Replace(Mid(DNStr, iPos + 1), "$", "")
      sTemp = Left(DNStr, iPos - 1)
      DNStr = ThisWorkbook.Worksheets(sTemp).Range(ss).Item(1, 1).Text
   End If
End Function

Function DNReal(s As String) As Single
   Dim ss As String
   ss = vMainWbk.Names(s)
   ss = Application.Substitute(ss, Chr(34), "")
   DNReal = Val(Mid(ss, 2, Len(ss) - 1))
End Function

Function AddressStr(Optional Row1 As Variant, Optional Col1 As Variant, Optional Row2 As Variant, Optional Col2 As Variant, _
         Optional R1C1 As Variant, Optional Wbk_Name As Variant, Optional Sht_Name As Variant) As String
   Dim s As String, u As Integer
   If IsMissing(Row2) Or IsMissing(Col2) Then
      Row2 = 0: Col2 = 0
   End If

   If IsMissing(Row1) Then Row1 = -1
   If IsMissing(R1C1) Then R1C1 = 0
   If R1C1 = xlR1C1 Then
      If Row2 >= 1 Then s = ":R" & LTrim(str(Row2)) & "C" & LTrim(str(Col2))
      s = "R" & LTrim(str(Row1)) & "C" & LTrim(str(Col1)) & s
   Else
      u = Int((Col1 - 1) / 26)
      If u >= 1 Then s = Chr(u + 64)
      s = s & Chr(((Col1 - 1) Mod 26) + 65) & LTrim(str(Row1))
      If Row2 >= 1 Then
         s = s & ":"
         u = Int((Col2 - 1) / 26)
         If u >= 1 Then s = s & Chr(u + 64)
         s = s & Chr(((Col2 - 1) Mod 26) + 65) & LTrim(str(Row2))
      End If
   End If
   s = Application.Substitute(s, "-1", "")
   If IsMissing(Sht_Name) = False Then
      If Len(Sht_Name) >= 1 Then
         s = Sht_Name & "'!" & s
         If IsMissing(Wbk_Name) = False Then
            If Len(Wbk_Name) > 0 Then s = "[" & Wbk_Name & "]" & s
         End If
         s = "'" & s
      End If
   End If
   AddressStr = s
End Function

  '---- Adds spaces to a string (filenames, Fieldnames, etc) that has only
    '     upper case characters to indicate where words start.
    '     For example: 'TheBoyRuns123' would be returned as the 'The Boy Runs 123'
    Function SpaceAdder(ByVal sName As String) As String
      Dim sChar As String, iChr As Integer, iChrs As Integer
      Dim iAsc, iLastAsc As Integer, sLastChr As String
      SpaceAdder = ""
      iChrs = Len(sName)
      iLastAsc = 120
      sLastChr = "a"
      For iChr = 1 To iChrs
        sChar = (Mid(sName, iChr, 1))
        If sChar = "_" Then sChar = " "
        iAsc = Asc(sChar)
        'If iAsc = 50 Then Stop
        If iAsc > 40 And iAsc <= 90 Or iAsc = 38 Then ' >"(" and <="Z" or ="&"
          If sLastChr = UCase(sLastChr) And sLastChr <> "&" Then
            SpaceAdder = SpaceAdder & sChar
          Else
            SpaceAdder = SpaceAdder & " " & sChar
            If sChar = "A" And iChr < iChrs Then
              sChar = Mid(sName, iChr + 1, 1)
              If sChar >= "A" And sChar <= "Z" Then
                SpaceAdder = SpaceAdder & " "
              End If
            End If
          End If
        Else
          If sLastChr = " " Then sChar = UCase(sChar)
          SpaceAdder = SpaceAdder & sChar
        End If
        iLastAsc = iAsc
        sLastChr = sChar
      Next iChr
      SpaceAdder = Trim(Replace(SpaceAdder, " ,", ", "))
      SpaceAdder = Trim(Replace(SpaceAdder, "  ", " "))
    End Function

Sub ApplyFormats()
    Dim vShtFormats As Worksheet
    Dim iFCol As Integer
    Dim sTemp As String
    Dim iFRow As Variant
    Range("3:3").WrapText = True
    Set vShtFormats = ThisWorkbook.Worksheets("Formats")
    
    iFCol = 1
    While Len(Cells(3, iFCol).Text) >= 2
       iFRow = Application.Match(Trim(Cells(3, iFCol).Text), vShtFormats.Range("B:B"), 0)
       If Not IsError(iFRow) Then
          With Cells(3, iFCol).EntireColumn
             sTemp = vShtFormats.Cells(iFRow, 5).Text 'Number Format
             .NumberFormat = sTemp
             
             sTemp = vShtFormats.Cells(iFRow, 6).Text 'Alignment
             If sTemp = "xlCenter" Then
                .HorizontalAlignment = xlCenter
             ElseIf sTemp = "xlLeft" Then
                .HorizontalAlignment = xlLeft
             ElseIf sTemp = "xlRight" Then
                .HorizontalAlignment = xlRight
             Else
             End If 'Alignment
             
             sTemp = vShtFormats.Cells(iFRow, 7).Text 'ColumnWidth
             If Len(sTemp) >= 1 Then
                .ColumnWidth = Val(sTemp)
                If Val(sTemp) >= 50 Then .WrapText = True
             End If
             
             sTemp = UCase(vShtFormats.Cells(iFRow, 8).Text) 'WrapText
             If sTemp = "YES" Then
                .WrapText = True
             End If
          
             sTemp = UCase(vShtFormats.Cells(iFRow, 9).Text) 'Hide?
             If sTemp = "YES" Then
                .Hidden = True
             End If
          
          End With  'EntireColumn
       Else
         ' MsgBox "Formats for Column '" & Cells(3, iFCol).Text & "' where not found.  Add this column to the ColumnFormats sheet.", vbCritical, "Formatting Problem"
       End If       'Format Found
       iFCol = iFCol + 1
    Wend            'Column Names
End Sub

Function FileNameSanitizer(ByVal sFileName As String) As String
Dim sFN_Start As String
sFN_Start = sFileName
FileNameSanitizer = Replace(sFileName, "*", "")
FileNameSanitizer = Replace(FileNameSanitizer, "<", "")
FileNameSanitizer = Replace(FileNameSanitizer, """", "'")
FileNameSanitizer = Replace(FileNameSanitizer, ">", "")
FileNameSanitizer = Replace(FileNameSanitizer, "?", "")
FileNameSanitizer = Replace(FileNameSanitizer, "[", "")
FileNameSanitizer = Replace(FileNameSanitizer, "]", "")
FileNameSanitizer = Replace(FileNameSanitizer, ":", "")
FileNameSanitizer = Replace(FileNameSanitizer, "|", "")
FileNameSanitizer = Replace(FileNameSanitizer, "\\", "\")
FileNameSanitizer = Left(FileNameSanitizer, 218)
If FileNameSanitizer <> sFN_Start Then
   'MsgBox "Warning: Filename of '" & sFN_Start & " ' had to be changed to '" & _
   FileNameSanitizer & "' to conform to Windows filenaming standards.", vbCritical
End If

End Function

Sub test()
Dim bTemp As Boolean
   sTemp = TextSanitizer("XSS/*-*/STYLE=xss:e/**/xpression(alert(097531))>", bTemp, True, True, True)


End Sub
Function TextSanitizer(ByVal sText As String, ByRef bChangesMade As Boolean, ByVal bFileNaming As Boolean, ByVal bSQL As Boolean, ByVal bScripting As Boolean) As String

Dim sStart As String
sStart = sText
'---- Script and Illegal Filename Characters -----
If bFileNaming Or bScripting Then
TextSanitizer = Replace(sStart, "*", "")
TextSanitizer = Replace(TextSanitizer, "<", "")
TextSanitizer = Replace(TextSanitizer, """", "'")
TextSanitizer = Replace(TextSanitizer, ">", "")
TextSanitizer = Replace(TextSanitizer, "?", "")
TextSanitizer = Replace(TextSanitizer, "[", "")
TextSanitizer = Replace(TextSanitizer, "]", "")
TextSanitizer = Replace(TextSanitizer, ":", "")
TextSanitizer = Replace(TextSanitizer, "|", "")
TextSanitizer = Replace(TextSanitizer, "\\", "\")
End If

If bFileNaming Then TextSanitizer = Left(TextSanitizer, 218)

If bScripting Or bSQL Then
   TextSanitizer = Replace(TextSanitizer, "'", " ")
   TextSanitizer = Replace(TextSanitizer, "/", " ")
End If

If bSQL Then
 TextSanitizer = TextSanitizer & " "
 If InStr(TextSanitizer, "select ") >= 1 Then
    TextSanitizer = Replace(TextSanitizer, Mid(TextSanitizer, InStr(sText, "select "), 6), " ")
 End If
 If InStr(TextSanitizer, "insert ") >= 1 Then
    TextSanitizer = Replace(TextSanitizer, Mid(TextSanitizer, InStr(sText, "select "), 6), " ")
 End If
 If InStr(TextSanitizer, "update ") >= 1 Then
    TextSanitizer = Replace(TextSanitizer, Mid(TextSanitizer, InStr(sText, "select "), 6), " ")
 End If
 If InStr(TextSanitizer, "delete ") >= 1 Then
    TextSanitizer = Replace(TextSanitizer, Mid(TextSanitizer, InStr(sText, "select "), 6), " ")
 End If
 TextSanitizer = Trim(TextSanitizer)
End If

bChangesMade = TextSanitizer <> sStart
'MsgBox "Warning: Filename of '" & sFN_Start & " ' had to be changed to '" & _
'TextSanitizer & "' to conform to Windows filenaming standards.", vbCritical

End Function



