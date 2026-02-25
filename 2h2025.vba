Sub ExportEmployeeForms_final()
    Dim wsRaw As Worksheet
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim employeeCode As String
    Dim employeeName As String
    Dim department As String
    Dim formType As String
    Dim targetCell As String
    Dim targetSheet As String
    Dim newWb As Workbook
    Dim wsTarget As Worksheet
    Dim savePath As String
    Dim fileName As String
    Dim exportCount As Integer
    Dim startTime As Double
    Dim logRow As Long
    Dim validForms As Variant
    Dim isValidForm As Boolean
    Dim j As Integer
    Dim fso As Object
    Dim tempPath As String
    Dim tempFile As String
    Dim sourceFileExt As String
    Dim pasteRanges As Variant
    Dim rng As Variant
    
    ' Bat dau dem thoi gian
    startTime = Timer
    exportCount = 0
    
    ' Danh sach cac form hop le
    validForms = Array("1", "2", "3", "4.1", "4.2", "5", "6", "7", "8", "9")
    
    ' Thiet lap duong dan luu file
    savePath = "D:\1. H-CONFIDENTAL\4.H - Special\2H2025\Export All\"
    
    ' Xac dinh dinh dang file goc
    sourceFileExt = LCase(Right(ThisWorkbook.Name, 4))
    If sourceFileExt <> ".xls" And sourceFileExt <> "xlsx" And sourceFileExt <> "xlsm" Then
        sourceFileExt = LCase(Right(ThisWorkbook.Name, 5))
    End If
    
    ' Tao duong dan file tam
    tempPath = Environ("TEMP")
    If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
    
    ' Tao thu muc su dung FileSystemObject
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(savePath) Then
        fso.CreateFolder savePath
    End If
    Set fso = Nothing
    On Error GoTo 0
    
    ' Thiet lap sheet du lieu nguon
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("Final_Consolidate_2H2025")
    On Error GoTo 0
    
    If wsRaw Is Nothing Then
        MsgBox "Khong tim thay sheet 'Final_Consolidate_2H2025'!", vbCritical
        Exit Sub
    End If
    
    ' Tao hoac xoa du lieu cu trong sheet Log
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Export_Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "Export_Log"
    Else
        wsLog.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Tao header cho Log
    With wsLog
        .Range("A1").Value = "STT"
        .Range("B1").Value = "Ma NV"
        .Range("C1").Value = "Ten NV"
        .Range("D1").Value = "Phong ban"
        .Range("E1").Value = "Loai Form"
        .Range("F1").Value = "Ten File"
        .Range("G1").Value = "Trang thai"
        .Range("H1").Value = "Thoi gian"
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(200, 200, 200)
    End With
    logRow = 2
    
    ' Tim dong cuoi cung co du lieu
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "E").End(xlUp).Row
    
    ' Tat cap nhat man hinh de tang toc do
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual ' Tat tinh toan tu dong
    
    ' Lap qua tung dong du lieu
    For i = 3 To lastRow
        ' Lay thong tin nhan vien
        employeeCode = Trim(wsRaw.Cells(i, "E").Value) ' Cot E
        employeeName = Trim(wsRaw.Cells(i, "F").Value) ' Cot F
        department = Trim(wsRaw.Cells(i, "G").Value) ' Cot G
        formType = Trim(wsRaw.Cells(i, "AF").Value) ' Cot AF
        
        ' Bo qua neu khong co ma nhan vien
        If employeeCode <> "" Then
            
            ' Kiem tra formType co nam trong danh sach hop le khong
            isValidForm = False
            For j = LBound(validForms) To UBound(validForms)
                If formType = validForms(j) Then
                    isValidForm = True
                    Exit For
                End If
            Next j
            
            ' Neu formType khong hop le, chi ghi log va bo qua
            If Not isValidForm Then
                wsLog.Cells(logRow, 1).Value = "-"
                wsLog.Cells(logRow, 2).Value = employeeCode
                wsLog.Cells(logRow, 3).Value = employeeName
                wsLog.Cells(logRow, 4).Value = department
                wsLog.Cells(logRow, 5).Value = formType
                wsLog.Cells(logRow, 6).Value = "N/A"
                wsLog.Cells(logRow, 7).Value = "Form type khong hop le - Khong tao file"
                wsLog.Cells(logRow, 8).Value = Now
                wsLog.Cells(logRow, 7).Interior.Color = RGB(255, 255, 200) ' Mau vang nhat
                logRow = logRow + 1
                GoTo NextIteration
            End If
            
            ' Xac dinh sheet dich va o dich
            targetSheet = formType
            
            ' Xac dinh o dich dua tren loai form
            If formType = "4.1" Or formType = "8" Then
                targetCell = "L6"
            Else
                targetCell = "M6"
            End If
            
            ' Kiem tra sheet co ton tai khong
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets(targetSheet)
            On Error GoTo 0
            
            If Not wsTarget Is Nothing Then
                ' Tao file tam voi dinh dang giong file goc
                tempFile = tempPath & "TempCopy_" & i & sourceFileExt
                
                ' Xoa file tam neu ton tai
                On Error Resume Next
                Kill tempFile
                On Error GoTo 0
                
                ' Luu file hien tai voi tat ca cac sheet
                On Error Resume Next
                ThisWorkbook.SaveCopyAs tempFile
                If Err.Number <> 0 Then
                    ' Ghi log loi khi tao file tam
                    wsLog.Cells(logRow, 1).Value = "-"
                    wsLog.Cells(logRow, 2).Value = employeeCode
                    wsLog.Cells(logRow, 3).Value = employeeName
                    wsLog.Cells(logRow, 4).Value = department
                    wsLog.Cells(logRow, 5).Value = formType
                    wsLog.Cells(logRow, 6).Value = "N/A"
                    wsLog.Cells(logRow, 7).Value = "Loi tao file tam: " & Err.Description
                    wsLog.Cells(logRow, 8).Value = Now
                    wsLog.Cells(logRow, 7).Interior.Color = RGB(255, 182, 193)
                    logRow = logRow + 1
                    Err.Clear
                    GoTo NextIteration
                End If
                On Error GoTo 0
                
                ' Mo file tam
                On Error Resume Next
                Set newWb = Workbooks.Open(fileName:=tempFile, ReadOnly:=False, UpdateLinks:=False)
                If Err.Number <> 0 Then
                    ' Ghi log loi khi mo file tam
                    wsLog.Cells(logRow, 1).Value = "-"
                    wsLog.Cells(logRow, 2).Value = employeeCode
                    wsLog.Cells(logRow, 3).Value = employeeName
                    wsLog.Cells(logRow, 4).Value = department
                    wsLog.Cells(logRow, 5).Value = formType
                    wsLog.Cells(logRow, 6).Value = "N/A"
                    wsLog.Cells(logRow, 7).Value = "Loi mo file tam: " & Err.Description
                    wsLog.Cells(logRow, 8).Value = Now
                    wsLog.Cells(logRow, 7).Interior.Color = RGB(255, 182, 193)
                    logRow = logRow + 1
                    Err.Clear
                    Kill tempFile
                    GoTo NextIteration
                End If
                On Error GoTo 0
                
                ' Gan ma nhan vien vao o dich
                newWb.Sheets(targetSheet).Range(targetCell).Value = employeeCode
                
                newWb.Sheets(targetSheet).Calculate  ' Tinh toan sheet
                Application.Calculate                 ' Tinh toan workbook
                DoEvents                              ' Cho Excel xu ly xong
                
                ' Paste value cac vung chua cong thuc tuy theo loai form
                On Error Resume Next
                If formType = "4.1" Or formType = "8" Then
                    ' Paste value cho form 4.1 va 8
                    pasteRanges = Array("C6:H8", "L6:N8", "C9:N10", "K20:K24", "M20:M24", "K39:K46", "M39:M46", "A51:N54", "A61:N66")
                Else
                    ' Paste value cho form 1,2,3,4.2,5,6,7,9
                    pasteRanges = Array("D6:I8", "D9:O10", "M5:O8", "L20:L24", "N20:N24", "L39:L44", "N39:N44", "B50:O50", "B52:O52", "B59:O64")
                End If
                
                For Each rng In pasteRanges
                    With newWb.Sheets(targetSheet).Range(rng)
                        .Value = .Value ' Paste value
                    End With
                Next rng
                On Error GoTo 0
                
                ' Xoa tat ca cac sheet khac, chi giu lai sheet can thiet
                Dim ws As Worksheet
                Dim wsToDelete As Collection
                Set wsToDelete = New Collection
                
                ' Thu thap danh sach sheet can xoa
                For Each ws In newWb.Worksheets
                    If ws.Name <> targetSheet Then
                        wsToDelete.Add ws.Name
                    End If
                Next ws
                
                ' Xoa cac sheet
                Dim wsName As Variant
                For Each wsName In wsToDelete
                    newWb.Sheets(CStr(wsName)).Delete
                Next wsName
                
                ' Doi ten sheet thanh ma nhan vien
                newWb.Sheets(targetSheet).Name = employeeCode
                
                ' Tao ten file
                fileName = department & "_" & employeeCode & "_" & employeeName & ".xlsx"
                ' Loai bo ky tu khong hop le trong ten file
                fileName = Replace(fileName, "/", "_")
                fileName = Replace(fileName, "\", "_")
                fileName = Replace(fileName, ":", "_")
                fileName = Replace(fileName, "*", "_")
                fileName = Replace(fileName, "?", "_")
                fileName = Replace(fileName, """", "_")
                fileName = Replace(fileName, "<", "_")
                fileName = Replace(fileName, ">", "_")
                fileName = Replace(fileName, "|", "_")
                
                ' Luu file moi (luon luu thanh .xlsx khong co macro)
                On Error Resume Next
                newWb.SaveAs fileName:=savePath & fileName, FileFormat:=xlOpenXMLWorkbook
                If Err.Number = 0 Then
                    ' Ghi log thanh cong
                    wsLog.Cells(logRow, 1).Value = exportCount + 1
                    wsLog.Cells(logRow, 2).Value = employeeCode
                    wsLog.Cells(logRow, 3).Value = employeeName
                    wsLog.Cells(logRow, 4).Value = department
                    wsLog.Cells(logRow, 5).Value = formType
                    wsLog.Cells(logRow, 6).Value = fileName
                    wsLog.Cells(logRow, 7).Value = "Thanh cong"
                    wsLog.Cells(logRow, 8).Value = Now
                    wsLog.Cells(logRow, 7).Interior.Color = RGB(144, 238, 144) ' Mau xanh la
                    exportCount = exportCount + 1
                Else
                    ' Ghi log loi
                    wsLog.Cells(logRow, 1).Value = "-"
                    wsLog.Cells(logRow, 2).Value = employeeCode
                    wsLog.Cells(logRow, 3).Value = employeeName
                    wsLog.Cells(logRow, 4).Value = department
                    wsLog.Cells(logRow, 5).Value = formType
                    wsLog.Cells(logRow, 6).Value = fileName
                    wsLog.Cells(logRow, 7).Value = "Loi luu file: " & Err.Description
                    wsLog.Cells(logRow, 8).Value = Now
                    wsLog.Cells(logRow, 7).Interior.Color = RGB(255, 182, 193) ' Mau do nhat
                End If
                On Error GoTo 0
                
                logRow = logRow + 1
                
                ' Dong workbook moi khong luu
                newWb.Close SaveChanges:=False
                
                ' Xoa file tam
                On Error Resume Next
                Kill tempFile
                On Error GoTo 0
                
            Else
                ' Ghi log neu khong tim thay sheet
                wsLog.Cells(logRow, 1).Value = "-"
                wsLog.Cells(logRow, 2).Value = employeeCode
                wsLog.Cells(logRow, 3).Value = employeeName
                wsLog.Cells(logRow, 4).Value = department
                wsLog.Cells(logRow, 5).Value = formType
                wsLog.Cells(logRow, 6).Value = "N/A"
                wsLog.Cells(logRow, 7).Value = "Khong tim thay sheet: " & formType
                wsLog.Cells(logRow, 8).Value = Now
                wsLog.Cells(logRow, 7).Interior.Color = RGB(255, 182, 193) ' Mau do nhat
                logRow = logRow + 1
            End If
            
            Set wsTarget = Nothing
        End If
        
NextIteration:
    Next i
    
    ' Format log sheet
    wsLog.Columns("A:H").AutoFit
    
    ' Bat lai cac thiet lap
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Thong bao hoan thanh
    MsgBox "Hoan thanh!" & vbCrLf & _
           "So file da tao: " & exportCount & vbCrLf & _
           "Tong so dong xu ly: " & (lastRow - 2) & vbCrLf & _
           "Thoi gian: " & Format((Timer - startTime) / 60, "0.00") & " phut" & vbCrLf & _
           "Duong dan: " & savePath, vbInformation, "Xuat file hoan tat"
    
    ' Kich hoat sheet Log
    wsLog.Activate
    
End Sub

