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
'��̏�
Const Adrs_ReceiptYear = "Z3"
Const Adrs_WorkDayReceipt = "K20"

Sub main()

    Dim arrUserDates() As Variant
    Dim arrDates() As Variant
    Dim wbDetail As Workbook
    Dim wbReceipt As Workbook
    Dim wbLunch As Workbook
    
    embedTemplate   '�ЂȌ`�ɔN��������͂���
    arrUserDates() = getUserDates() '���O���擾

    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���~
    With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .EnableEvents = False
    End With

    '���׏��쐬����
    Workbooks.Add '���[�N�u�b�N���쐬(���׏��p)
    Set wbDetail = ActiveWorkbook
    
    wbCreateDetails arrUserDates, wbDetail
    wbDetail.SaveAs ThisWorkbook.Path & "\���׏�" & ".xlsx" 'TODO:timestamp
    wbDetail.Close
    
    '��̏��쐬����
    Workbooks.Add
    Set wbReceipt = ActiveWorkbook
    
    wbCreateReceipt arrUserDates, wbReceipt
    wbReceipt.SaveAs ThisWorkbook.Path & "\��̏�" & ".xlsx"
    wbReceipt.Close
    
    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���ĊJ
    With Application
      .Calculation = xlCalculationAutomatic
      .ScreenUpdating = True
      .EnableEvents = True
    End With

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
    
    '��̏��̓��t������
    With ThisWorkbook.Worksheets("Receipt")
        .Range(Adrs_ReceiptYear).Value = Year(dayPayed)
    End With
    
End Sub

Sub wbCreateDetails(arrUserDates As Variant, wbDetail As Workbook)
    
    Dim wsDetail As Worksheet
    Dim i As Long
    
    For i = LBound(arrUserDates, 1) To UBound(arrUserDates, 1)
            ThisWorkbook.Worksheets("Detail").Copy after:=wbDetail.ActiveSheet '�V�K���[�N�u�b�N��sheet1�̑O
            Set wsDetail = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
            wsDetail.Name = arrUserDates(i, Name)
            wsDetail.Range(Adrs_WorkDayReceipt).Value = arrUserDates(i, daysWorked)
    Next
    
    Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    wbDetail.Sheets("sheet1").Delete
    wbDetail.Sheets(1).Select
    Application.DisplayAlerts = True
    
End Sub

Sub wbCreateReceipt(arrUserDates As Variant, wbReceipt As Workbook)

    Dim wsReceipt As Worksheet
    Dim i As Long
    
    For i = LBound(arrUserDates, 1) To UBound(arrUserDates, 1)
            ThisWorkbook.Worksheets("Receipt").Copy after:=wbReceipt.ActiveSheet '�V�K���[�N�u�b�N��sheet1�̑O
            Set wsReceipt = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
            wsReceipt.Name = arrUserDates(i, Name)
            wsReceipt.Range(Adrs_WorkDayReceipt).Value = arrUserDates(i, daysWorked)
    Next
    
    Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    wbReceipt.Sheets("sheet1").Delete
    wbReceipt.Sheets(1).Select
    Application.DisplayAlerts = True
    
End Sub

Function getUserDates() As Variant

    Dim returnArray As Variant, n As Long
    n = Cells(Rows.Count, 1).End(xlUp).Row - 4
    ReDim returnArray(n, 11)
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A5", Cells(Rows.Count, 11).End(xlUp)).Value
'    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1:K20").Value
    getUserDates = returnArray
End Function


