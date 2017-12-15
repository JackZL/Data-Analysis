# This samples is used for showing that VBA is able to manipulate EXCEL, WORD, EMAIL in one place
Sub MakeScar()

    Dim row1 As Integer
    Dim supplierCode As Long
    Dim array1()
    Dim int1 As Integer
    Dim myWordApp As Word.Application
    Dim myDoc As Word.Document
    Dim myWorkbook As Workbook
    Dim myapp As Excel.Application
    Dim strName As String
    Dim mySheet As Worksheet
    Dim sheetPath As String
    Dim row2 As Integer
    Dim serialNo As String
    Dim wk2 As Object
    
    Application.ScreenUpdating = False
    'Set sheetPath = ActiveWorkbook.Path
    row1 = Selection.Row
    If InStr(Cells(row1, 10), "-") > 0 Then
        supplierCode = Fix(left(Cells(row1, 10), 6))
    Else
        supplierCode = Fix(Cells(row1, 10))
    End If
    Cells(row1, 10).Value = supplierCode
    Columns("L:N").NumberFormat = "General"
    
    ' find supplier contact information from sheet "approved vendor list"
    Cells(row1, 12).Formula = _
    "=VLOOKUP(" & supplierCode & ",'C:\Users\CNJALIU11\Documents\10 Supplier audit\[Approved Vendor List-2016.09.01.xlsx]ForContact'!$A$1:$Q$400,9,FALSE)"
    Cells(row1, 13).Formula = _
    "=VLOOKUP(" & supplierCode & ",'C:\Users\CNJALIU11\Documents\10 Supplier audit\[Approved Vendor List-2016.09.01.xlsx]ForContact'!$A$1:$Q$400,10,FALSE)"
    Cells(row1, 14).Formula = _
    "=VLOOKUP(" & supplierCode & ",'C:\Users\CNJALIU11\Documents\10 Supplier audit\[Approved Vendor List-2016.09.01.xlsx]ForContact'!$A$1:$Q$400,11,FALSE)"
    range(Cells(row1, 12), Cells(row1, 14)).Copy
    range(Cells(row1, 12), Cells(row1, 14)).PasteSpecial (xlPasteValues)
    array1() = range(Cells(row1, 1), Cells(row1, 150))
    
    '获取SCAR List文件
'    Set myWorkbook = GetObject("C:\Users\CNJALIU11\Documents\13 Supplier improvement\SCAR list.xlsx")
'    Set mySheet = myWorkbook.Worksheets("SCAR Detail")
    Set myWorkbook = Workbooks.Open("C:\Users\CNJALIU11\Documents\13 Supplier improvement\SCAR list.xlsx")
    Set mySheet = myWorkbook.Worksheets("SCAR Detail")

    With mySheet
        'find out the last row
        row2 = .range("a1").End(xlDown).Row
        ' set serial no as YYYYMMDD_145
        .Cells(row2 + 1, 1) = Format(Date, "YYYYMMDD") & "_" & CStr(row2 + 1)
        strName = Format(Date, "YYYYMMDD") & "_" & CStr(row2 + 1) & "_" & array1(1, 2) & "_" & array1(1, 10) & ".docx"
        .Cells(row2 + 1, 2) = array1(1, 2)
        .Cells(row2 + 1, 3) = array1(1, 5)
        .Cells(row2 + 1, 4) = Date
        .Cells(row2 + 1, 5) = array1(1, 11)
        .Cells(row2 + 1, 6) = array1(1, 77)
        .Cells(row2 + 1, 11) = strName
    End With
    'Windows(myWorkbook.name).Visible = True
    '如果使用getobject()，上一句必须添加，否则打不开
    myWorkbook.Close (True)
    'myWorkbook.Save
    Set myWorkbook = Nothing
    
    '创建SCAR文件
    Set myWordApp = CreateObject("word.application")
    Set myDoc = myWordApp.Documents.Add("C:\Users\CNJALIU11\Documents\Custom Office Templates\QSF 195 SCAR rev.B.dotx")
    
    With myDoc.Tables(1).Columns(2)
        .Cells(1).range.Text = array1(1, 2)
        .Cells(2).range.Text = array1(1, 3)
        .Cells(3).range.Text = array1(1, 5)
        .Cells(4).range.Text = Date
        .Cells(8).range.Text = array1(1, 10) & " " & array1(1, 11)
        .Cells(9).range.Text = array1(1, 12)
        .Cells(10).range.Text = array1(1, 13) & " " & array1(1, 14)
    End With
    
    myDoc.SaveAs2 ("C:\Users\CNJALIU11\Desktop\" & strName)
    myDoc.Application.Quit
    
    Set myDoc = Nothing
  
    
    '创建邮件
    Dim mailApp As Outlook.Application
    Dim mail As Outlook.MailItem
    
    Set mailApp = CreateObject("Outlook.Application")
    Set mail = mailApp.CreateItem(olMailItem)
    
    On Error Resume Next
    With mail
            .To = array1(1, 14)
            .CC = "alan-lianzhen.wang@cn.abb.com"
            .subject = strName & " SCAR"
            .BodyFormat = olFormatHTML
            .HTMLBody = "<html><head><style type = 'text/css'>p {font-family: arial; font-size:14pt;}</style></head><body><p>Hi  " & array1(1, 12) _
                & "<p>See attachment. <br>Material: " & array1(1, 2) & "<br>Quantity: " & array1(1, 111) & "<br>Please find out the root cause and corrective action of mentioned quality issue." _
                & "<br>Please fill out the attachment and finish the SCAR within <strong>1 weeks</strong>. <br><br>Best Regards<br>Jack Zhen Liu" _
                & "<br>Supplier Quality Engineer<br>Cell Phone: +86 181 1617 6797</p></body></html>"

            .Attachments.Add ("C:\Users\CNJALIU11\Desktop\" & strName)
            .Display
    End With
    
    Set mailApp = Nothing
    Set mail = Nothing
    'Set arraya1 = Nothing
    
    Application.ScreenUpdating = True

End Sub
