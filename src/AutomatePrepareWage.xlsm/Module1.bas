Attribute VB_Name = "Module1"
Option Explicit
Sub main()
'TODO
'Workbook���쐬���ϐ��ɃZ�b�g
'Workbook��ҏW(���O�ύXetc)
    '�ЂȌ`�쐬�������Ăяo��
        '
    Dim wbDetail As Workbook
    Dim arr() As Variant
    
    arr() = getUserNames()
    
    Workbooks.Add '�V�K���[�N�u�b�N���쐬
    Set wbDetail = ActiveWorkbook
    
    Dim val As Variant
    For Each val In arr
        wbCreateDetails val, wbDetail
    Next
    
    wbDetail.SaveAs "���׏�.xlsx"
    wbDetail.Close

End Sub

Sub wbCreateDetails(val As Variant, wbDetail As Workbook)

    Dim wsDetail As Worksheet

    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1) '�V�K���[�N�u�b�N��sheet1�̑O

    Set wsDetail = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
    wsDetail.Cells(1, 1).Value = val
    
End Sub

Function getUserNames() As Variant

    Dim returnArray As Variant
'    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1", Cells(Rows.Count, 1).End(xlUp)).Value
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1:A20").Value
    getUserNames = returnArray
End Function
