Attribute VB_Name = "Module24"
' =====================================================================================
' Procedure : ClosingStockValuation_FIFO_CodeByCyril()
' Author    : Cyril
' Version   : v3.0
' Date      : [Insert Date]
'
' Purpose   : Automates Closing Stock Valuation using FIFO method.
'
' Description:
'   - Reads data from a Purchase Register and Closing Stock sheet.
'   - Applies FIFO (First In, First Out) logic to value closing stock.
'   - Generates:
'       * A detailed Closing Stock Valuation report.
'       * A summary sheet with subtotals and grand total.
'
' Inputs:
'   - Purchase Register sheet (prompted from user).
'   - Closing Stock sheet (prompted from user).
'
' Outputs:
'   - "ClosingStockValuation" worksheet (Detailed Report).
'   - "SummaryReport" worksheet (Summary with totals).
'
' Notes:
'   - Report and summary sheets are cleared/created fresh each run.
'   - Includes error handling for invalid data or missing sheets.
'
'
' License   : MIT License (see LICENSE file in repository)
' GitHub    : https://github.com/codebycyril/fifo-stock-valuation
' =====================================================================================

Sub ClosingStockValuation_FIFO_CodeByCyril()
' (Original name retained for reference: ClosingStockValuation_FIFO_CodeByCyril())

    Dim wsPurchase As Worksheet
    Dim wsClosing As Worksheet
    Dim wsReport As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRowPurchase As Long
    Dim lastRowClosing As Long
    Dim reportRow As Long
    Dim summaryRow As Long
    Dim i As Long, j As Long
    Dim material As String
    Dim closingQuantity As Double
    Dim purchaseQuantity As Double
    Dim amount As Double
    Dim rate As Double
    Dim remainingQty As Double
    Dim totalValueToStock As Double
    Dim cumulativeQty As Double
    Dim purchaseSheetName As String
    Dim closingSheetName As String
    Dim reportSheetName As String
    Dim summarySheetName As String
    Dim foundMatch As Boolean
    Dim productDescription As String
    Dim lastMaterial As String

    On Error GoTo ErrorHandler

    ' ---------------------------------------------------------------------
    ' Prompt user for sheet names
    ' ---------------------------------------------------------------------
    purchaseSheetName = InputBox("Enter the name of the Purchase Register sheet:", "Sheet Name Input", "PurchaseRegister")
    closingSheetName = InputBox("Enter the name of the Closing Stock sheet:", "Sheet Name Input", "ClosingStock")
    reportSheetName = "ClosingStockValuation"
    summarySheetName = "SummaryReport"

    ' ---------------------------------------------------------------------
    ' Ensure the worksheets exist
    ' ---------------------------------------------------------------------
    Set wsPurchase = ThisWorkbook.Sheets(purchaseSheetName)
    Set wsClosing = ThisWorkbook.Sheets(closingSheetName)
    
    ' Create or clear the detailed report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets(reportSheetName)
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = reportSheetName
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Create or clear the summary report sheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets(summarySheetName)
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add
        wsSummary.Name = summarySheetName
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' =========================
    ' Setup Detailed Report with heading + headers
    ' =========================
    With wsReport
        .Cells(1, 1).Value = "Closing Stock Valuation (Detailed Report)"
        .Range("A1:K1").Merge
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "Product Reference"
        .Cells(2, 2).Value = "Product Description"
        .Cells(2, 3).Value = "Posting Date"
        .Cells(2, 4).Value = "Document Date"
        .Cells(2, 5).Value = "Bill Number"
        .Cells(2, 6).Value = "Vendor Name"
        .Cells(2, 7).Value = "Quantity in Bill"
        .Cells(2, 8).Value = "Subtotal"
        .Cells(2, 9).Value = "Qty to Stock"
        .Cells(2, 10).Value = "Value to Stock"
        .Cells(2, 11).Value = "Reference"
        .Rows(2).Font.Bold = True
    End With

    ' =========================
    ' Setup Summary Report with heading + headers
    ' =========================
    With wsSummary
        .Cells(1, 1).Value = "Closing Stock Valuation Summary"
        .Range("A1:E1").Merge
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "Product Reference"
        .Cells(2, 2).Value = "Product Description"
        .Cells(2, 3).Value = "Quantity"
        .Cells(2, 4).Value = "Rate"
        .Cells(2, 5).Value = "Subtotal"
        .Rows(2).Font.Bold = True
    End With

    ' Start data from row 3
    reportRow = 3
    summaryRow = 3

    ' ---------------------------------------------------------------------
    ' Find the last rows with data in each worksheet
    ' ---------------------------------------------------------------------
    lastRowPurchase = wsPurchase.Cells(wsPurchase.Rows.Count, "A").End(xlUp).Row
    lastRowClosing = wsClosing.Cells(wsClosing.Rows.Count, "A").End(xlUp).Row

    ' ---------------------------------------------------------------------
    ' Sort purchase register by Posting Date to apply FIFO correctly
    ' ---------------------------------------------------------------------
    wsPurchase.Sort.SortFields.Clear
    wsPurchase.Sort.SortFields.Add Key:=wsPurchase.Range("C2:C" & lastRowPurchase), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With wsPurchase.Sort
        .SetRange wsPurchase.Range("A1:I" & lastRowPurchase)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' ---------------------------------------------------------------------
    ' Initialize variables
    ' ---------------------------------------------------------------------
    lastMaterial = ""

    ' ---------------------------------------------------------------------
    ' Loop through closing stock data
    ' ---------------------------------------------------------------------
    For i = 2 To lastRowClosing
        material = wsClosing.Cells(i, 1).Value
        closingQuantity = wsClosing.Cells(i, 2).Value
        productDescription = ""
        cumulativeQty = 0
        totalValueToStock = 0
        foundMatch = False

        ' Loop through purchase register data
        For j = 2 To lastRowPurchase
            If wsPurchase.Cells(j, 1).Value = material Then
                If productDescription = "" Then
                    productDescription = wsPurchase.Cells(j, 2).Value
                End If
                purchaseQuantity = wsPurchase.Cells(j, 7).Value
                amount = wsPurchase.Cells(j, 8).Value

                If IsNumeric(purchaseQuantity) And IsNumeric(amount) Then
                    rate = amount / purchaseQuantity
                Else
                    MsgBox "Non-numeric data found in Quantity or Subtotal columns. Please check your data.", vbCritical
                    Exit Sub
                End If

                foundMatch = True

                If cumulativeQty < closingQuantity Then
                    If cumulativeQty + purchaseQuantity >= closingQuantity Then
                        remainingQty = closingQuantity - cumulativeQty
                        wsReport.Cells(reportRow, 9).Value = remainingQty
                        wsReport.Cells(reportRow, 10).Value = (remainingQty / purchaseQuantity) * amount
                        With wsReport
                            .Cells(reportRow, 1).Value = material
                            .Cells(reportRow, 2).Value = productDescription
                            .Cells(reportRow, 3).Value = wsPurchase.Cells(j, 3).Value
                            .Cells(reportRow, 4).Value = wsPurchase.Cells(j, 4).Value
                            .Cells(reportRow, 5).Value = wsPurchase.Cells(j, 5).Value
                            .Cells(reportRow, 6).Value = wsPurchase.Cells(j, 6).Value
                            .Cells(reportRow, 7).Value = purchaseQuantity
                            .Cells(reportRow, 8).Value = amount
                            .Cells(reportRow, 11).Value = wsPurchase.Cells(j, 9).Value
                        End With
                        totalValueToStock = totalValueToStock + (remainingQty / purchaseQuantity) * amount
                        cumulativeQty = closingQuantity
                        reportRow = reportRow + 1
                        Exit For
                    Else
                        wsReport.Cells(reportRow, 9).Value = purchaseQuantity
                        wsReport.Cells(reportRow, 10).Value = amount
                        With wsReport
                            .Cells(reportRow, 1).Value = material
                            .Cells(reportRow, 2).Value = productDescription
                            .Cells(reportRow, 3).Value = wsPurchase.Cells(j, 3).Value
                            .Cells(reportRow, 4).Value = wsPurchase.Cells(j, 4).Value
                            .Cells(reportRow, 5).Value = wsPurchase.Cells(j, 5).Value
                            .Cells(reportRow, 6).Value = wsPurchase.Cells(j, 6).Value
                            .Cells(reportRow, 7).Value = purchaseQuantity
                            .Cells(reportRow, 8).Value = amount
                            .Cells(reportRow, 11).Value = wsPurchase.Cells(j, 9).Value
                        End With
                        totalValueToStock = totalValueToStock + amount
                        cumulativeQty = cumulativeQty + purchaseQuantity
                        reportRow = reportRow + 1
                    End If
                End If
            End If
        Next j

        ' Case: no matching data in Purchase Register
        If Not foundMatch Then
            productDescription = wsClosing.Cells(i, 3).Value
            With wsReport
                .Cells(reportRow, 1).Value = material
                .Cells(reportRow, 2).Value = productDescription
                .Cells(reportRow, 3).Value = "Opening Stock"
                .Cells(reportRow, 5).Value = "No data in Purchase Register"
                .Cells(reportRow, 6).Value = 0
                .Cells(reportRow, 7).Value = 0
                .Cells(reportRow, 8).Value = closingQuantity
                .Cells(reportRow, 9).Value = closingQuantity
                .Cells(reportRow, 10).Value = 0
            End With
            reportRow = reportRow + 1
        End If

        ' Totals line in detailed report
        If totalValueToStock > 0 Then
            With wsReport
                .Cells(reportRow, 1).Value = material
                .Cells(reportRow, 2).Value = productDescription
                .Cells(reportRow, 9).Value = closingQuantity
                .Cells(reportRow, 10).Value = totalValueToStock
                .Cells(reportRow, 9).Font.Bold = True
                .Cells(reportRow, 10).Font.Bold = True
                .Cells(reportRow, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Cells(reportRow, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
                .Cells(reportRow, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Cells(reportRow, 10).Borders(xlEdgeBottom).LineStyle = xlDouble
            End With
            reportRow = reportRow + 1
        End If

        ' Blank row for visual separation
        reportRow = reportRow + 1

        ' Write to summary report with reference to detailed report
        With wsSummary
            .Cells(summaryRow, 1).Value = material
            .Cells(summaryRow, 2).Value = productDescription
            .Cells(summaryRow, 3).Value = closingQuantity
            .Cells(summaryRow, 4).Value = rate
            .Cells(summaryRow, 5).Formula = "='" & reportSheetName & "'!J" & reportRow - 2
        End With
        summaryRow = summaryRow + 1

        lastMaterial = material
    Next i

    ' ---------------------------------------------------------------------
    ' Final formatting
    ' ---------------------------------------------------------------------
    wsReport.Columns.AutoFit
    wsReport.Columns("F:J").NumberFormat = "#,##0.00"
    wsSummary.Columns.AutoFit
    wsSummary.Columns("C:E").NumberFormat = "#,##0.00"

    ' ---------------------------------------------------------------------
    ' Add grand total line in Summary
    ' ---------------------------------------------------------------------
    Dim lastDataRow As Long
    lastDataRow = summaryRow - 1
    With wsSummary
        .Cells(lastDataRow + 1, 1).Value = "Total Stock Value"
        .Range(.Cells(lastDataRow + 1, 1), .Cells(lastDataRow + 1, 4)).Merge
        .Cells(lastDataRow + 1, 5).Formula = "=SUM(E3:E" & lastDataRow & ")"
        With .Range(.Cells(lastDataRow + 1, 1), .Cells(lastDataRow + 1, 5))
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
            .Borders.LineStyle = xlContinuous
        End With
    End With

    MsgBox "ClosingStockValuation_FIFO_CodeByCyril: Completed Successfully.", vbInformation
    Exit Sub

' -------------------------------------------------------------------------
' Error handler
' -------------------------------------------------------------------------
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub


