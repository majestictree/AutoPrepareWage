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
'���H
Const Adrs_DateOfIssue = "D1"
Const Adrs_NumOfLunch = "Q14"
Const Adrs_LunchThisJPYear = "B19"
Const Adrs_LunchLastMonth = "D19"
Const Adrs_LunchReceiptJPYear = "F23"

Sub main()

    Dim arrUserDates() As Variant
    Dim arrDates() As Variant
    Dim wbDetail As Workbook
    Dim wbReceipt As Workbook
    Dim wbLunch As Workbook
    Dim tailOfFileName As String
    Dim dayPayed As Date
    Dim lastMonthDate As Date
    
    embedTemplate   '�ЂȌ`�ɔN��������͂���
    arrUserDates() = getUserDates() '���f�[�^��z��֊i�[���鏈���֐�
    
    '�ۑ��t�@�C�����̖�����������
    dayPayed = ThisWorkbook.Worksheets("InvoiceData").Range("B1").Value
    lastMonthDate = DateSerial(Year(dayPayed), Month(dayPayed) - 1, 1)
    tailOfFileName = "(" & Format(lastMonthDate, "yyyymm") & ").xlsx"
    
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
    wbDetail.SaveAs ThisWorkbook.Path & "\�H���x�����׏�" & tailOfFileName
    wbDetail.Close
    
    '��̏��쐬����
    Workbooks.Add
    Set wbReceipt = ActiveWorkbook
    
    wbCreateReceipt arrUserDates, wbReceipt
    wbReceipt.SaveAs ThisWorkbook.Path & "\�H����̏�" & tailOfFileName
    wbReceipt.Close
    
    '���H�㐿�������̎���
    Workbooks.Add
    Set wbLunch = ActiveWorkbook
    
    wbCreateLunch arrUserDates, wbLunch
    wbLunch.SaveAs ThisWorkbook.Path & "\���H�㐿�������̎���" & tailOfFileName
    wbLunch.Close
    
    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���ĊJ
    With Application
      .Calculation = xlCalculationAutomatic
      .ScreenUpdating = True
      .EnableEvents = True
    End With

End Sub

Sub embedTemplate()

    Dim dayPayed As Date
    Dim lastMonthDate As Date

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
    
    '���H�̓��t������
    With ThisWorkbook.Worksheets("Lunch")
        .Range(Adrs_DateOfIssue).Value = dayPayed
        .Range(Adrs_LunchThisJPYear).Value = Year(lastMonthDate) - 2018
        .Range(Adrs_LunchLastMonth).Value = Month(lastMonthDate)
        .Range(Adrs_LunchReceiptJPYear).Value = Year(dayPayed) - 2018
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

Sub wbCreateLunch(arrUserDates As Variant, wbLunch As Workbook)
    
    Dim wsLunch As Worksheet
    Dim i As Long
    
    For i = LBound(arrUserDates, 1) To UBound(arrUserDates, 1)
            ThisWorkbook.Worksheets("Lunch").Copy after:=wbLunch.ActiveSheet '�V�K���[�N�u�b�N��sheet1�̑O
            Set wsLunch = ActiveSheet '�R�s�[�����V�[�g��ϐ��ɃZ�b�g
            wsLunch.Name = arrUserDates(i, Name)
            wsLunch.Range(Adrs_NumOfLunch).Value = arrUserDates(i, nLunch)
    Next
    
    Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    wbLunch.Sheets("sheet1").Delete
    wbLunch.Sheets(1).Select
    Application.DisplayAlerts = True
    
End Sub


Function getUserDates() As Variant

    Dim returnArray As Variant, n As Long
    n = Cells(Rows.Count, 1).End(xlUp).Row - 4
    ReDim returnArray(n, 11)
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A5", Cells(Rows.Count, 11).End(xlUp)).Value
    getUserDates = returnArray
End Function


