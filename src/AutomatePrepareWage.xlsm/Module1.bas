Attribute VB_Name = "Module1"
Option Explicit
Sub main()
'TODO
'Workbookを作成し変数にセット
'Workbookを編集(名前変更etc)
    'ひな形作成処理を呼び出し
        '


End Sub

Sub wbCreateDetails()

    Dim wsDetail As Worksheet
    Dim wbDetail As Workbook
    
    Workbooks.Add '新規ワークブックを作成
    Set wbDetail = ActiveWorkbook
    
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1) '新規ワークブックのsheet1の前
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1)
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1)

'    Set wsDetail = ActiveSheet 'コピーしたシートを変数にセット
'    wsDetail.Cells(1, 1).Value = "test"
    
    wbDetail.SaveAs "明細書.xlsx"
    wbDetail.Close

End Sub
