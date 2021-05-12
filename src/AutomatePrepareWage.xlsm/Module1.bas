Attribute VB_Name = "Module1"
Option Explicit

'�񖼂̗񋓑�
Enum colUserDate
    Name = 1
    daysWorked
    unit
    teate
    baseWage
    addWage
    totalHere
    nLunch
    selfPay
    totalAll
End Enum

'�l��}������ʒu��萔�\�L
'���׏�
Const Adrs_WorkDay = "F8"
Const Adrs_nLunch = "E17"
Const Adrs_LastJPYear = "F3"
Const Adrs_PayJPYear = "F4"
Const Adrs_LastMonth = "H3"
Const Adrs_PayMonth = "H4"
Const Adrs_PayDay = "J4"

Sub main()

    Dim arrUserDates() As Variant
    Dim arrDates() As Variant
    Dim wbDetail As Workbook
    Dim wbReceipt As Workbook
    Dim wbLunch As Workbook
    
    '�ЂȌ`�ɔN��������͂���
    embedTemplate
    
    arrUserDates() = getUserDates() '���O���擾
'    arrDates() = getInvoiceDates()

    '���׏��쐬����
    Workbooks.Add '���[�N�u�b�N���쐬(���׏��p)
    Set wbDetail = ActiveWorkbook
    
    wbCreateDetails arrUserDates, wbDetail
    wbDetail.SaveAs "���׏�.xlsx" 'TODO:�ۑ�����w��
    wbDetail.Close
    

End Sub

Sub embedTemplate()

    Dim dayPayed As Date
    Dim lastMonthDate As Long

    dayPayed = ThisWorkbook.Worksheets("InvoiceData").Range("B1").Value
    lastMonthDate = DateSerial(Year(dayPayed), Month(dayPayed) - 1, 1)
    '���׏��̓��t������
    With ThisWorkbook.Worksheets("Detail")
        .Range(Adrs_LastJPYear).Value = Year(lastMonthDate) - 2018
        .Range(Adrs_PayJPYear).Value = Year(dayPayed) - 2018
        .Range(Adrs_LastMonth).Value = Month(lastMonthDate)
        .Range(Adrs_PayMonth).Value = Month(dayPayed)
        .Range(Adrs_PayDay).Value = Day(dayPayed)
    End With
    
End Sub

Sub wbCreateDetails(arrUserDates As Variant, wbDetail As Workbook)
    
    Dim wsDetail As Worksheet
    Dim i As Long
    
    For i = LBound(arrUserDates, 1) To UBound(arrUserDates, 1)
            ThisWorkbook.Worksheets("Detail").Copy after:=wbDetail.ActiveSheet '�V�K���[�N�u�b�N��sheet1�̑O
            Set wsDetail = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
            wsDetail.Name = arrUserDates(i, Name)
            wsDetail.Range(Adrs_WorkDay).Value = arrUserDates(i, daysWorked)
            wsDetail.Range(Adrs_nLunch).Value = arrUserDates(i, nLunch)
    Next
    
    Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    wbDetail.Sheets("sheet1").Delete
    Application.DisplayAlerts = True
    
End Sub

Function getUserDates() As Variant

    Dim returnArray As Variant
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A5", Cells(Rows.Count, 11).End(xlUp)).Value
'    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1:K20").Value
    getUserDates = returnArray
End Function


