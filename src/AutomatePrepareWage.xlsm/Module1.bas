Attribute VB_Name = "Module1"
Option Explicit
Sub main()
'TODO
'Workbook���쐬���ϐ��ɃZ�b�g
'Workbook��ҏW(���O�ύXetc)
    '�ЂȌ`�쐬�������Ăяo��
        '


End Sub

Sub wbCreateDetails()

    Dim wsDetail As Worksheet
    Dim wbDetail As Workbook
    
    Workbooks.Add '�V�K���[�N�u�b�N���쐬
    Set wbDetail = ActiveWorkbook
    
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1) '�V�K���[�N�u�b�N��sheet1�̑O
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1)
    ThisWorkbook.Worksheets("Detail").Copy before:=wbDetail.Sheets(1)

'    Set wsDetail = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
'    wsDetail.Cells(1, 1).Value = "test"
    
    wbDetail.SaveAs "���׏�.xlsx"
    wbDetail.Close

End Sub
