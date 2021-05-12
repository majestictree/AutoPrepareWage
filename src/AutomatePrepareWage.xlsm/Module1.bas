Attribute VB_Name = "Module1"
Option Explicit
Sub main()
'TODO
'Workbookを作成し変数にセット
'Workbookを編集(名前変更etc)
    'ひな形作成処理を呼び出し
        '
    Dim wbDetail As Workbook
    Dim arr() As Variant
    
    arr() = getUserNames()
    
    Workbooks.Add '新規ワークブックを作成
    Set wbDetail = ActiveWorkbook
    
    Dim val As Variant
    For Each val In arr
        wbCreateDetails val, wbDetail
    Next
    
    wbDetail.SaveAs "明細書.xlsx"
    wbDetail.Close

End Sub

Sub wbCreateDetails(val As Variant, wbDetail As Workbook)

    Dim wsDetail As Worksheet

    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1) '新規ワークブックのsheet1の前

    Set wsDetail = ActiveSheet 'コピーしたシートを変数にセット
    wsDetail.Cells(1, 1).Value = val
    
End Sub

Function getUserNames() As Variant

    Dim returnArray As Variant
'    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1", Cells(Rows.Count, 1).End(xlUp)).Value
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1:A20").Value
    getUserNames = returnArray
End Function
