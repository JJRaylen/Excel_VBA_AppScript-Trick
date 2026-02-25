Sub CreateSymmetricalTable_WithValidation_FINAL()
    ' ============= BOOST HIEU SUAT VOI FILE MILLION ROWS =============
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' =================================================
    
    
    Dim wsSource As Worksheet, wsResult As Worksheet, wsError As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long, errorRow As Long
    Dim voucher As String, currentVoucher As String
    Dim negAccounts As Object, posAccounts As Object
    Dim negKey As Variant, posKey As Variant
    Dim totalNeg As Double, totalPos As Double
    
    ' ===== SETUP =====
    Set wsSource = ThisWorkbook.Sheets("Source")
    
    ' Tao sheet ket qua
    Set wsResult = ThisWorkbook.Sheets.Add
    wsResult.Name = "Result_" & Format(Now, "hhmmss")
    wsResult.Range("A1:D1").Value = Array("Voucher", "Negative", "Positive", "Value")
    wsResult.Range("A1:D1").Font.Bold = True
    wsResult.Range("A1:D1").Interior.Color = RGB(144, 238, 144) ' Xanh lá
    
    ' Tao sheet loi
    Set wsError = ThisWorkbook.Sheets.Add
    wsError.Name = "Errors_" & Format(Now, "hhmmss")
    wsError.Range("A1:E1").Value = Array("Voucher", "Account", "Type", "Remaining", "Issue")
    wsError.Range("A1:E1").Font.Bold = True
    wsError.Range("A1:E1").Interior.Color = RGB(255, 200, 200) ' Ð? nh?t
    
    outputRow = 2
    errorRow = 2
    
    ' ===== SAP XEP DU LIEU =====
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    With wsSource.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsSource.Range("A2:A" & lastRow), Order:=xlAscending
        .SetRange wsSource.Range("A1:D" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    ' ===== KHOI TAO DICTIONARIES =====
    Set negAccounts = CreateObject("Scripting.Dictionary")
    Set posAccounts = CreateObject("Scripting.Dictionary")
    currentVoucher = ""
    
    ' ===== XU LY TONG VOUCHER =====
    For i = 2 To lastRow
        voucher = wsSource.Cells(i, 1).Value
        
        ' Khi gAp voucher moi, xu lý voucher cu
        If voucher <> currentVoucher And currentVoucher <> "" Then
            Call ProcessVoucherWithValidation(currentVoucher, negAccounts, posAccounts, _
                                              wsResult, outputRow, wsError, errorRow)
            negAccounts.RemoveAll
            posAccounts.RemoveAll
        End If
        
        currentVoucher = voucher
        
        ' Thu thap data
        If wsSource.Cells(i, 4).Value = "Negative" Then
            negAccounts(wsSource.Cells(i, 2).Value) = Abs(wsSource.Cells(i, 3).Value)
        Else
            posAccounts(wsSource.Cells(i, 2).Value) = wsSource.Cells(i, 3).Value
        End If
    Next i
    
    ' Xu lý voucher cuoi
    If negAccounts.count > 0 Then
        Call ProcessVoucherWithValidation(currentVoucher, negAccounts, posAccounts, _
                                          wsResult, outputRow, wsError, errorRow)
    End If
    
    ' ===== FORMAT VÀ THÔNG BÁO =====
    wsResult.Columns("A:D").AutoFit
    wsError.Columns("A:E").AutoFit
    
    If errorRow = 2 Then
        Application.DisplayAlerts = False
        wsError.Delete
        Application.DisplayAlerts = True
        MsgBox "Completed! No Error Found." & vbCrLf & _
               "Result: " & wsResult.Name, vbInformation
    Else
        MsgBox "Completed with " & (errorRow - 2) & " error" & vbCrLf & _
               "Result: " & wsResult.Name & vbCrLf & _
               "Error Infor: " & wsError.Name, vbExclamation
    End If
End Sub

Sub ProcessVoucherWithValidation(voucher As String, _
                                  negAccounts As Object, _
                                  posAccounts As Object, _
                                  wsResult As Worksheet, _
                                  ByRef outputRow As Long, _
                                  wsError As Worksheet, _
                                  ByRef errorRow As Long)
    Dim negKey As Variant, posKey As Variant
    Dim negValue As Double, posValue As Double
    Dim allocatedValue As Double
    Dim totalNegBefore As Double, totalPosBefore As Double
    Dim tolerance As Double
    
    tolerance = 0.01 ' Sai s? ch?p nh?n du?c (1 xu)
    
    ' ===== TÍNH T?NG BAN Ð?U =====
    totalNegBefore = 0
    For Each negKey In negAccounts.Keys
        totalNegBefore = totalNegBefore + negAccounts(negKey)
    Next negKey
    
    totalPosBefore = 0
    For Each posKey In posAccounts.Keys
        totalPosBefore = totalPosBefore + posAccounts(posKey)
    Next posKey
    
    ' ===== KIEM TRA BALANCE TONG THE =====
    If Abs(totalNegBefore - totalPosBefore) > tolerance Then
        wsError.Cells(errorRow, 1).Value = voucher
        wsError.Cells(errorRow, 2).Value = "ALL"
        wsError.Cells(errorRow, 3).Value = "VOUCHER"
        wsError.Cells(errorRow, 4).Value = totalNegBefore - totalPosBefore
        wsError.Cells(errorRow, 5).Value = "Sum Negative <> Sum Positive"
        wsError.Cells(errorRow, 4).Font.Color = RGB(255, 0, 0)
        errorRow = errorRow + 1
    End If
    
    ' ===== GREEDY MATCHING =====
    For Each negKey In negAccounts.Keys
        negValue = negAccounts(negKey)
        
        ' Duy?t qua các Positive accounts
        For Each posKey In posAccounts.Keys
            If posAccounts(posKey) > tolerance And negValue > tolerance Then
                
                ' Tính allocation
                allocatedValue = Application.WorksheetFunction.Min(negValue, posAccounts(posKey))
                
                ' Ghi output
                wsResult.Cells(outputRow, 1).Value = voucher
                wsResult.Cells(outputRow, 2).Value = negKey
                wsResult.Cells(outputRow, 3).Value = posKey
                wsResult.Cells(outputRow, 4).Value = Round(allocatedValue, 2)
                outputRow = outputRow + 1
                
                ' C?p nh?t giá tr? còn l?i
                negValue = negValue - allocatedValue
                posAccounts(posKey) = posAccounts(posKey) - allocatedValue
                
                If negValue < tolerance Then Exit For
            End If
        Next posKey
        
        ' ===== KIEM TRA NEGATIVE ACCOUNT CÒN DU =====
        If negValue > tolerance Then
            wsError.Cells(errorRow, 1).Value = voucher
            wsError.Cells(errorRow, 2).Value = negKey
            wsError.Cells(errorRow, 3).Value = "NEGATIVE"
            wsError.Cells(errorRow, 4).Value = Round(negValue, 2)
            wsError.Cells(errorRow, 5).Value = "Khong du Positive de match"
            wsError.Cells(errorRow, 4).Font.Color = RGB(255, 0, 0)
            errorRow = errorRow + 1
        End If
    Next negKey
    
    ' ===== KIEM TRA POSITIVE ACCOUNTS CÒN DU =====
    For Each posKey In posAccounts.Keys
        If posAccounts(posKey) > tolerance Then
            wsError.Cells(errorRow, 1).Value = voucher
            wsError.Cells(errorRow, 2).Value = posKey
            wsError.Cells(errorRow, 3).Value = "POSITIVE"
            wsError.Cells(errorRow, 4).Value = Round(posAccounts(posKey), 2)
            wsError.Cells(errorRow, 5).Value = "Không du Negative de match"
            wsError.Cells(errorRow, 4).Font.Color = RGB(255, 0, 0)
            errorRow = errorRow + 1
        End If
    Next posKey
    
        ' ============= END CODE HO TRO BOOST HIEU SUAT =============
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    
End Sub

