Attribute VB_Name = "Module1"
Option Explicit

'列名の列挙体
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

'値を挿入する位置を定数表記
'明細書
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
    
    'ひな形に年月日を入力する
    embedTemplate
    
    arrUserDates() = getUserDates() '名前を取得
'    arrDates() = getInvoiceDates()

    '明細書作成処理
    Workbooks.Add 'ワークブックを作成(明細書用)
    Set wbDetail = ActiveWorkbook
    
    wbCreateDetails arrUserDates, wbDetail
    wbDetail.SaveAs "明細書.xlsx" 'TODO:保存先を指定
    wbDetail.Close
    

End Sub

Sub embedTemplate()

    Dim dayPayed As Date
    Dim lastMonthDate As Long

    dayPayed = ThisWorkbook.Worksheets("InvoiceData").Range("B1").Value
    lastMonthDate = DateSerial(Year(dayPayed), Month(dayPayed) - 1, 1)
    '明細書の日付等入力
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
            ThisWorkbook.Worksheets("Detail").Copy after:=wbDetail.ActiveSheet '新規ワークブックのsheet1の前
            Set wsDetail = ActiveSheet 'コピーしたシートを変数にセット
            wsDetail.Name = arrUserDates(i, Name)
            wsDetail.Range(Adrs_WorkDay).Value = arrUserDates(i, daysWorked)
            wsDetail.Range(Adrs_nLunch).Value = arrUserDates(i, nLunch)
    Next
    
    Application.DisplayAlerts = False ' メッセージを非表示
    wbDetail.Sheets("sheet1").Delete
    Application.DisplayAlerts = True
    
End Sub

Function getUserDates() As Variant

    Dim returnArray As Variant
    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A5", Cells(Rows.Count, 11).End(xlUp)).Value
'    returnArray = ThisWorkbook.Worksheets("InvoiceData").Range("A1:K20").Value
    getUserDates = returnArray
End Function


