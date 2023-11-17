Attribute VB_Name = "Module1"
Public Function funCalculate(rngYear As String)
Dim wsYear As Worksheet
Dim wsCalculate As Worksheet
Dim uniqueRng As Range
Dim myCol As Long
Dim row As Long
Dim myInputRng As Range
Dim myOutputRng As Range
Dim iForLoop As Long
Dim TotalRow As Long
Dim OpenPrice As Double
Dim ClosedPrice As Double
Dim tickerValue As String
    
    Application.ScreenUpdating = False
    
    Set wsYear = ThisWorkbook.Worksheets(rngYear)
    Set wsCalculate = ThisWorkbook.Worksheets("Calculate")
    
    Application.StatusBar = "Stock analysis is in Progress for " & rngYear
    
    With wsYear
        .Activate
    Set myOutputRng = .Range("I2")
    myOutputRng.Select
    
    '-------------------------
    TotalRow = ActiveSheet.UsedRange.Rows.Count
        .Range("I2:Q" & TotalRow).ClearContents
        .Range("I2:Q" & TotalRow).ClearFormats
        
        .Range("I1").value = "Ticker"
        .Range("J1").value = "Yearly Change"
        .Range("K1").value = "Percent Change"
        .Range("L1").value = "Total Stock Volume"
        .Range("P1").value = "Ticker"
        .Range("Q1").value = "Value"
        .Range("O2").value = "Greatest % Increase"
        .Range("O3").value = "Greatest % Decrease"
        .Range("O4").value = "Greatest Total Volume"
        
    
    row = .Cells(Rows.Count, "A").End(xlUp).row
    Set myInputRng = .Range("A2:A" & row)
    
    myInputRng.Select
    Selection.Copy
    .Activate
    myOutputRng.Select
    ActiveSheet.Paste
    
    TotalRow = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).row
    Application.CutCopyMode = False
    ActiveSheet.Range("I1:I" & TotalRow).RemoveDuplicates columns:=1, Header:= _
        xlNo
    '-------------------------
    
    TotalRow = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).row
    For iForLoop = 2 To TotalRow
    
        tickerValue = .Cells(iForLoop, 9).value
        
        'getting the opening and closed price for a ticker
        .Activate
        .Range("A1").Select

        Selection.AutoFilter
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        ActiveSheet.Range("$A$1:$G$" & row).AutoFilter Field:=1, Criteria1:=tickerValue
        ActiveSheet.Range("$A$1:$G$" & row).AutoFilter Field:=2, Criteria1:="=" & rngYear & "0102"
        OpenPrice = Format(ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 3).value, "#.00")
        
        ActiveSheet.ShowAllData
        
        ActiveSheet.Range("$A$1:$G$" & row).AutoFilter Field:=1, Criteria1:=tickerValue
        ActiveSheet.Range("$A$1:$G$" & row).AutoFilter Field:=2, Criteria1:="=" & rngYear & "1231"
        ClosedPrice = Format(ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6).value, "#.00")
        
        ActiveSheet.ShowAllData
        
        .Cells(iForLoop, 10).value = Format(ClosedPrice - OpenPrice, "#.00")
        If .Cells(iForLoop, 10).value > 0 Then
            .Cells(iForLoop, 10).Interior.ColorIndex = 4
        Else
            .Cells(iForLoop, 10).Interior.ColorIndex = 3
        End If
        
        .Cells(iForLoop, 11).value = Format(.Cells(iForLoop, 10).value / OpenPrice, "#.00%")
        
        'Calculating the total stock volume of the stock
        .Cells(iForLoop, 12).value = Application.WorksheetFunction.SumIf(wsYear.Range("A:A"), tickerValue, wsYear.Range("G:G"))
    Next
    
    .Range("P2").value = Application.WorksheetFunction.Index(.Range("I:I").value, Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(.Range("K:K")), .Range("K:K"), 0), 1)
    .Range("P3").value = Application.WorksheetFunction.Index(.Range("I:I").value, Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(.Range("K:K")), .Range("K:K"), 0), 1)
    .Range("P4").value = Application.WorksheetFunction.Index(.Range("I:I").value, Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(.Range("L:L")), .Range("L:L"), 0), 1)
    .Range("Q2").value = Format(Application.WorksheetFunction.Max(.Range("K:K")), "#.00%")
    .Range("Q3").value = Format(Application.WorksheetFunction.Min(.Range("K:K")), "#.00%")
    .Range("Q4").value = Format(Application.WorksheetFunction.Max(.Range("L:L")), "#.00")
    
    End With
    
    wsCalculate.Activate
    wsCalculate.Range("A1").Select
    Application.StatusBar = "Stock analysis has been completed for " & rngYear
    
    Application.ScreenUpdating = True
    
End Function

Public Sub finalMicroRun()
Dim varYear As String

varYear = Worksheets("Calculate").Range("rngYear")

Application.ScreenUpdating = False
Application.StatusBar = "Stock analysis has started for year: " & varYear
Select Case varYear
    Case "2018"
        funCalculate ("2018")
    Case "2019"
        funCalculate ("2019")
    Case "2020"
        funCalculate ("2020")
    Case Default
        MsgBox "Invalid Year, please select a valid value frpm the list", vbInformation, "Message"
    End Select

Application.ScreenUpdating = True
End Sub




