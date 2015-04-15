Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public gBaseLines As Integer
Public intS1 As Integer
Public intS2 As Integer

Dim SourceBook As Workbook
Dim TargetBook(2) As Workbook


Sub Pivot_HZ()
Attribute Pivot_HZ.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim strActiveSH As String
    Dim strSheetName As String
    Dim strDataSheet As String
    Dim strPivotName As String

    strDataSheet = "HZ"
    strPivotName = "Pivot2"
    strSheetName = "ERP-" & strDataSheet
    
    Sheets.Add
    strActiveSH = ActiveSheet.Name
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=strDataSheet & "!A:R", Version:=xlPivotTableVersion10) _
        .CreatePivotTable TableDestination:=strActiveSH & "!R3C1", TableName:=strPivotName, DefaultVersion:=xlPivotTableVersion10
    Sheets(strActiveSH).Select
    Cells(3, 1).Select
    
'    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'    ActiveChart.SetSourceData Source:=Range(ActiveSH & "!$A$3:$G$16")
'    ActiveChart.Parent.Delete

    With ActiveSheet.PivotTables(strPivotName).PivotFields("ETA")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(strPivotName).PivotFields("Material")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveWindow.SmallScroll Down:=6
    ActiveSheet.PivotTables(strPivotName).AddDataField ActiveSheet.PivotTables(strPivotName).PivotFields("Order Quantity"), "加總 - Order Quantity", xlSum
    With ActiveSheet.PivotTables(strPivotName).PivotFields("ETA")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    Rows("4:4").Select:     Selection.NumberFormatLocal = "yyyy/mm/dd"
    Columns("A:A").Select: Selection.NumberFormatLocal = "@"

    ActiveSheet.Name = strSheetName
    Cells(1, 1).Select
End Sub

Sub Pivot_SZ()
    Dim strActiveSH As String
    Dim strSheetName As String
    Dim strDataSheet As String
    Dim strPivotName As String

    strDataSheet = "SZ"
    strPivotName = "Pivot1"
    strSheetName = "ERP-" & strDataSheet
    
    Sheets.Add
    strActiveSH = ActiveSheet.Name
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=strDataSheet & "!A:R", Version:=xlPivotTableVersion10) _
        .CreatePivotTable TableDestination:=strActiveSH & "!R3C1", TableName:=strPivotName, DefaultVersion:=xlPivotTableVersion10
    Sheets(strActiveSH).Select
    Cells(3, 1).Select
    
'    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'    ActiveChart.SetSourceData Source:=Range(ActiveSH & "!$A$3:$G$16")
'    ActiveChart.Parent.Delete

    With ActiveSheet.PivotTables(strPivotName).PivotFields("ETA")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(strPivotName).PivotFields("Material")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ActiveWindow.SmallScroll Down:=6
    ActiveSheet.PivotTables(strPivotName).AddDataField ActiveSheet.PivotTables(strPivotName).PivotFields("Order Quantity"), "加總 - Order Quantity", xlSum
    With ActiveSheet.PivotTables(strPivotName).PivotFields("ETA")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    Rows("4:4").Select:     Selection.NumberFormatLocal = "yyyy/mm/dd"
    Columns("A:A").Select: Selection.NumberFormatLocal = "@"
    
    ActiveSheet.Name = strSheetName
    Cells(1, 1).Select
End Sub

Sub 清除資料_Click()
    Application.DisplayAlerts = False   ' turn off the screen updating

' 清除 總表  -----------------------------
    Sheets("總表").Select
    
'    Range("A:P").Select
'    Selection.ClearContents
'    With Selection.Interior     ' 清除儲存格 填滿格式
'        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

    Cells.Select:    Selection.Delete Shift:=xlUp
    
    Cells(2, "A").Select
    
    ' 清除 SZ  -----------------------------
    Sheets("SZ").Select
    Cells.Select:    Selection.Delete Shift:=xlUp

    Cells(2, "A").Select

    ' 清除 HZ  -----------------------------
    Sheets("HZ").Select
    Cells.Select:    Selection.Delete Shift:=xlUp

    Cells(2, "A").Select

    ' 清除 樞紐分析表  -----------------------------
    Sheets("ERP-HZ").Select:    ActiveWindow.SelectedSheets.Delete
    Sheets("ERP-SZ").Select:    ActiveWindow.SelectedSheets.Delete

    Application.DisplayAlerts = True    ' turn on the screen updating
    
    ' 切換到 Menu 工作表
    Sheets("Menu").Select
    
    MsgBox "完成 資料清除作業", vbInformation
End Sub

Sub 讀取客戶資料_Click()
    Dim strFile As String
    Dim SourceWB As String
    Dim TargetWB As String
    Dim intDataRows As Long
    
    TargetWB = ActiveWorkbook.Name
'    TargetSH = ActiveSheet.Name
'
    strFile = Sheets("Menu").Cells(1, "B") & "\" & Sheets("Menu").Cells(2, "B")

    Application.ScreenUpdating = False  ' turn off the screen updating
    Application.DisplayAlerts = False

    Set SourceBook = Workbooks.Open(strFile, UpdateLinks:=True, ReadOnly:=True)         ' open the source workbook, read only
    SourceWB = ActiveWorkbook.Name

    Application.DisplayAlerts = False
    Windows(SourceWB).Activate
    Columns("A:P").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(TargetWB).Activate
    Sheets("總表").Select
    Range("A1").Select
'    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False     ' 選擇性貼上
    
    Selection.End(xlDown).Select
    intDataRows = ActiveCell().Row
    Range("P2").Select
    
    SourceBook.Close False              ' close the source workbook without saving any changes
    Set SourceBook = Nothing            ' free memory

'    Application.DisplayAlerts = True
    
    ' 設定 Q2 公式 =DATE(2015,1,MID(D2,1,2)*7-9)
    Range("Q1") = "ETA"
'    Range("Q2").Select
'    ActiveCell.FormulaR1C1 = "=DATE(2015,1,MID(RC[-10],1,2)*7-9)"
    Cells(2, "Q").Formula = "=DATE(P2,1,MID(D2,1,2)*7-9)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & intDataRows)
'    Range("Q2:Q774").Select
    
    ' 設定 R2 公式  =VLOOKUP(C2, 'D:\GitHub\_Misc\A180\[對照表_A180.xls]對照表'!$A:$B, 2, FALSE)
    Range("R1") = "Class"
    Cells(2, "R").Formula = "=VLOOKUP(C2, '" & ActiveWorkbook.Path & "\[對照表_A180.xls]對照表'!$A:$B, 2, FALSE)"
    Range("R2").Select
    Selection.AutoFill Destination:=Range("R2:R" & intDataRows)
   
    
    ' 刪除多餘的資料 ---------------------------------------------
    Selection.AutoFilter        ' 啟動 自動篩選 Menu
    
    ' 1.
'    ActiveSheet.Range("$A$1:$R$778").AutoFilter Field:=17, Criteria1:="<2015/3/24", Operator:=xlAnd
    ActiveSheet.Range("$A:$R").AutoFilter Field:=17, Criteria1:="<" & Sheets("Menu").Cells(6, "B"), Operator:=xlAnd
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
'    ActiveSheet.Range("$A$1:$R$742").AutoFilter Field:=17
    ActiveSheet.Range("$A:$R").AutoFilter Field:=17         '清除 篩選條件
   
    ' 2.
'    Selection.AutoFilter
    ActiveSheet.Range("$A:$R").AutoFilter Field:=17, Criteria1:=">" & Sheets("Menu").Cells(8, "B"), Operator:=xlAnd
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
'    ActiveSheet.Range("$A$1:$R$335").AutoFilter Field:=17
    ActiveSheet.Range("$A:$R").AutoFilter Field:=17         '清除 篩選條件
    
    Selection.AutoFilter        ' 關閉 自動篩選 Menu
    Range("A2").Select
    ' 刪除多餘的資料 ---------------------------------------------
    
    
    ' 填寫資料、製作 Pivot =======================
    Call WriteData_SZ
    Call Pivot_SZ

    Call WriteData_HZ
    Call Pivot_HZ
    ' 填寫資料、製作 Pivot =======================
    
   
    Sheets("總表").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("A1").Select
   
    Sheets("Menu").Select
    Range("A1").Select
    
    Application.ScreenUpdating = True  ' turn on the screen updating
    Application.DisplayAlerts = True
    
    MsgBox "完成 讀取客戶資料", vbInformation
End Sub

Sub WriteDatas(pTargetBook, pPos As Integer, pAry() As String)
    MsgBox pAry(1) & " - " & pAry(3) & " - " & pAry(7)
End Sub

Sub WriteData_SZ()
    Dim intDataRows As Long

    Sheets("總表").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    intDataRows = ActiveCell().Row

    Selection.AutoFilter
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 3
    ActiveSheet.Range("$A$1:$R$" & intDataRows).AutoFilter Field:=18, Criteria1:="SZ"
    Cells.Select
'    Range("C1").Activate
    Selection.Copy
    Sheets("SZ").Select
    Range("A1").Select
'    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False     ' 選擇性貼上
    
    Range("A1").Select
End Sub

Sub WriteData_HZ()
    Dim intDataRows As Long

    Sheets("總表").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    intDataRows = ActiveCell().Row

    Selection.AutoFilter
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 3
    ActiveSheet.Range("$A$1:$R$" & intDataRows).AutoFilter Field:=18, Criteria1:="HZ"
    Cells.Select
'    Range("C1").Activate
    Selection.Copy
    Sheets("HZ").Select
    Range("A1").Select
'    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False     ' 選擇性貼上

    Range("A1").Select
End Sub

'Sub CreateMail()
'    Dim objOutlook As Object
'    Dim objMail As Object
'    Dim rngTo As Range
'    Dim rngSubject As Range
'    Dim rngBody As Range
'    Dim rngAttach As Range
'
'    Set objOutlook = CreateObject("Outlook.Application")
'    Set objMail = objOutlook.CreateItem(0)
'
'    With ActiveSheet
'        Set rngTo = .Range("B1")
'        Set rngSubject = .Range("B2")
'        Set rngBody = .Range("B3")
'        Set rngAttach = .Range("B4")
'    End With
'
'    With objMail
'        .To = rngTo.Value
'        .Subject = rngSubject.Value
'        .Body = rngBody.Value
'        .Attachments.Add rngAttach.Value
'        .Display 'Instead of .Display, you can use .Send to send the email _
'                    or .Save to save a copy in the drafts folder
'    End With
'
'    Set objOutlook = Nothing
'    Set objMail = Nothing
'    Set rngTo = Nothing
'    Set rngSubject = Nothing
'    Set rngBody = Nothing
'    Set rngAttach = Nothing
'
'End Sub

'Sub AddAttachment()
' Dim myItem As Outlook.MailItem
' Dim myAttachments As Outlook.Attachments
'
' Set myItem = Application.CreateItem(olMailItem)
'' Set myAttachments = myItem.Attachments
' objMail.Attachments.Add "D:\Documents\Q496.xlsx", _
' olByValue, 1, "4th Quarter 1996 Results Chart"
' myItem.Display
'End Sub

Sub Mail2SZ_Click()
    Dim objOutlook As Object
    Dim objMail As Object
    Dim strTarget As String
    Dim strTargetFile As String
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    objMail.To = Cells(13, "A")
    If Cells(15, "A") <> "" Then
        objMail.Cc = Cells(15, "A")
    End If
    If Cells(17, "A") <> "" Then
        objMail.Bcc = Cells(17, "A")
    End If
    objMail.Subject = Cells(10, "B")
    objMail.Body = GetMailContent()
    
    If UCase(Right(Cells(2, "B"), 4)) = ".XLS" Then
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 14), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    Else
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 15), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    End If
    strTargetFile = ActiveWorkbook.Path & "\" & strTarget & ".xls"
    
    objMail.attachments.Add strTargetFile
'    objMail.Attachments.Add "D:\GitHub\Videos.mdb"
'    objMail.Attachments.Add Cells(1, "B") & "\" & Cells(2, "B")
    
    objMail.Display     ' .Display, you can use
'    objMail.Save        ' to save a copy in the drafts folder
'    objMail.Send        ' to send the email

    Set objOutlook = Nothing
    Set objMail = Nothing
End Sub

Sub Mail2HZ_Click()
    Dim objOutlook As Object
    Dim objMail As Object
    Dim strTarget As String
    Dim strTargetFile As String
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    objMail.To = Cells(21, "A")
    If Cells(23, "A") <> "" Then
        objMail.Cc = Cells(23, "A")
    End If
    If Cells(25, "A") <> "" Then
        objMail.Bcc = Cells(25, "A")
    End If
    objMail.Subject = Cells(10, "B")
    objMail.Body = GetMailContent()
    
    If UCase(Right(Cells(2, "B"), 4)) = ".XLS" Then
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 14), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    Else
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 15), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    End If
    strTargetFile = ActiveWorkbook.Path & "\" & strTarget & ".xls"
    
    objMail.attachments.Add strTargetFile
    
'    objMail.Attachments.Add "D:\GitHub\Videos.mdb"
'    objMail.Attachments.Add Cells(1, "B") & "\" & Cells(2, "B")
    
    objMail.Display     ' .Display, you can use
'    objMail.Save        ' to save a copy in the drafts folder
'    objMail.Send        ' to send the email

    Set objOutlook = Nothing
    Set objMail = Nothing
End Sub

Function GetMailContent()
    Dim strContent As String
    Dim cRow As Integer
    
    strContent = ""
    cRow = 12
    
    Do While (Cells(cRow, "B") <> "Cindy")
        strContent = strContent + Cells(cRow, "B") + vbCrLf
    
        cRow = cRow + 1
    Loop
    
    strContent = strContent + "Cindy     " + Format(Date, "yyyy/mm/dd")
    GetMailContent = strContent
End Function

Sub 轉成檔案()
    Call 轉換成EXCEL_SZ
    Call 轉換成EXCEL_HZ
End Sub

Sub 轉換成EXCEL_SZ()
    Dim NewBook As Workbook
    Dim SourceWB As String
    Dim SourceSH As String
    Dim TargetWB As String
    Dim strTarget As String
    Dim strTargetFile As String
    
    SourceWB = ActiveWorkbook.Name
    SourceSH = ActiveSheet.Name
    
    Set NewBook = Workbooks.Add
    TargetWB = NewBook.Name
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate
    Sheets("SZ").Select
    Cells.Select
    Selection.Copy
    Windows(TargetWB).Activate
'    Sheets.Add After:=ActiveSheet
'    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False     ' 選擇性貼上
    Sheets(1).Select        '"工作表1"
    Sheets(1).Name = "SZ"
    Range("A1").Select
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate
    Sheets("ERP-SZ").Select
    Cells.Select
    Range("A7").Activate
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows(TargetWB).Activate
    Sheets.Add After:=ActiveSheet
'    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False     ' 選擇性貼上
    Sheets(2).Select        '"工作表2"
    Sheets(2).Name = "ERP-SZ"
    Range("A1").Select
    
    Rows("4:4").Select:     Selection.NumberFormatLocal = "yyyy/mm/dd"
    Columns("A:A").Select: Selection.NumberFormatLocal = "@"
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate

    If UCase(Right(Cells(2, "B"), 4)) = ".XLS" Then
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 14), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    Else
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 15), 10) & ")_SZ"     ' A180 forecast (02-09-15)_SZ
    End If
    strTargetFile = ActiveWorkbook.Path & "\" & strTarget & ".xls"
    NewBook.SaveAs strTargetFile, Excel.XlFileFormat.xlExcel5
    
    NewBook.Close

    MsgBox "完成 " & strTarget & ".xls" & " 檔案建制工作", vbOKOnly
End Sub

Sub 轉換成EXCEL_HZ()
    Dim NewBook As Workbook
    Dim SourceWB As String
    Dim SourceSH As String
    Dim TargetWB As String
    Dim strTarget As String
    Dim strTargetFile As String
    
    SourceWB = ActiveWorkbook.Name
    SourceSH = ActiveSheet.Name
    
    Set NewBook = Workbooks.Add
    TargetWB = NewBook.Name
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate
    Sheets("HZ").Select
    Cells.Select
    Selection.Copy
    Windows(TargetWB).Activate
'    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets(1).Select        '"工作表1"
    Sheets(1).Name = "HZ"
    Range("A1").Select
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate
    Sheets("ERP-HZ").Select
    Cells.Select
    Range("A7").Activate
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows(TargetWB).Activate
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets(2).Select        '"工作表2"
    Sheets(2).Name = "ERP-HZ"
    Range("A1").Select
    
    Workbooks(SourceWB).Sheets(SourceSH).Activate

    If UCase(Right(Cells(2, "B"), 4)) = ".XLS" Then
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 14), 10) & ")_HZ"     ' A180 forecast (02-09-15)_HZ
    Else
        strTarget = "A180 forecast (" & Left(Right(Cells(2, "B"), 15), 10) & ")_HZ"     ' A180 forecast (02-09-15)_HZ
    End If
    strTargetFile = ActiveWorkbook.Path & "\" & strTarget & ".xls"
    NewBook.SaveAs strTargetFile, Excel.XlFileFormat.xlExcel5
    
    NewBook.Close

    MsgBox "完成 " & strTarget & ".xls" & " 檔案建制工作", vbOKOnly
End Sub
