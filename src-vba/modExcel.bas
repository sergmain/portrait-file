Attribute VB_Name = "modExcel"
Option Explicit

Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const constMaxCol = 50
Private Const xlCellTypeLastCell = 11
Private Const xlCellTypeFormulas = -4123
Private Const xlCellTypeConstants = 2
Private Const xlSheetVisible = -1
Private Const xlA1 = 1

Public m_oExcel As Object
Public m_oExcelMacro As Object

Public m_bExceled As Boolean


Public Function GetExcelDocContent1(oFileName As String, ByRef errorMessage As String, ByRef hideListsExist As Boolean) As Object
    Dim strVersion As String
    Dim excelWorkBook As Object, excelSheets As Object, excelSheet As Object, excelRng As Object
    Dim rng As Object
    Dim startRow As Long, startCol As Long
    Dim endRow As Long, endCol As Long
    Dim bInserted As Boolean
    Dim dispAlerts As Boolean
    Dim referenceStyle As Long
    Dim wordVersion As Long, contextRunned As Boolean
    Dim iSheet As Long
    
    On Error GoTo ErrLabel
    Set GetExcelDocContent1 = Nothing
    If Not m_oExcel Is Nothing Then
        strVersion = m_oExcel.Version
        If strVersion = "" Then
            Set m_oExcel = Nothing
        End If
    End If
    
    If m_oExcel Is Nothing Then
        Set m_oExcel = CreateObject("Excel.Application")
        If m_oExcel Is Nothing Then Exit Function
        m_oExcel.Visible = False
    End If
    
    dispAlerts = m_oExcel.DisplayAlerts
    m_oExcel.DisplayAlerts = False
    hideListsExist = False
    
    Set excelWorkBook = m_oExcel.WorkBooks.Open(fileName:=oFileName, ReadOnly:=True)
    If excelWorkBook Is Nothing Then
        errorMessage = "Ошибка при открытии"
        Exit Function
    End If
    
    
    wordVersion = GetWordVersion()
    contextRunned = ContextIsRun()
    
    Set excelSheets = excelWorkBook.Sheets
    If excelSheets Is Nothing Then Exit Function

    referenceStyle = m_oExcel.referenceStyle
    m_oExcel.referenceStyle = xlA1
    Set GetExcelDocContent1 = m_WordApp.Documents.Add(, 0, , False)
    GetExcelDocContent1.ActiveWindow.Visible = False
    Call OpenClipboard(0&)
    Call EmptyClipboard
    Call CloseClipboard

    For iSheet = 1 To excelSheets.Count
        Set excelSheet = excelSheets.Item(iSheet)
        If excelSheet.Visible = xlSheetVisible Or m_bHideLists Then
            Call GetExcelBorders(excelSheet, startRow, startCol, endRow, endCol)
            
            If endCol > 0 Then
                Call GetExcelDocContent1.Range.InsertParagraphAfter
                Set rng = GetExcelDocContent1.Range
                
                Set excelRng = excelSheet.Range(excelSheet.Cells(startRow, startCol), excelSheet.Cells(endRow, endCol))
                Call excelRng.Copy
                
                If wordVersion > 11 And contextRunned Then Call Sleep(3000)
                Call GetExcelDocContent1.Range(rng.End - 1, rng.End).PasteExcelTable(False, False, False)
                Call OpenClipboard(0&)
                Call EmptyClipboard
                Call CloseClipboard
                bInserted = True
                
                If excelSheet.Visible <> xlSheetVisible Then hideListsExist = True
            End If
        End If
    Next
    Call excelWorkBook.Close(False)
    
    If Not bInserted Then
        errorMessage = "Пустой файл"
        GoTo ErrLabel
    End If
    
    m_oExcel.DisplayAlerts = dispAlerts
    m_oExcel.referenceStyle = referenceStyle
    Exit Function
    
ErrLabel:
    m_oExcel.referenceStyle = referenceStyle
    If errorMessage = "" Then errorMessage = err.Description
    Call EmptyClipboard
    Call CloseClipboard
    On Error Resume Next
    If Not GetExcelDocContent1 Is Nothing Then
        Call GetExcelDocContent1.Close(False)
        Set GetExcelDocContent1 = Nothing
    End If
End Function

Public Sub GetExcelBorders(curSheet As Object, ByRef startRow As Long, ByRef startCol As Long, ByRef endRow As Long, ByRef endCol As Long)
    Dim printArea As String, subStr As String
    Dim num As Long
    Dim endColPrint As Long, startColPrint As Long
    Dim endRowPrint As Long, startRowPrint As Long

    startRow = 1:   endRow = curSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Call GetExcelLastCol(curSheet, startCol, endCol)
    printArea = curSheet.PageSetup.printArea
    
    If printArea <> "" Then
        printArea = Mid(printArea, 2)
        
        num = InStr(1, printArea, "$")
        subStr = Left(printArea, num - 1)
        printArea = Mid(printArea, num + 1)
        startColPrint = StringColumnToLong(subStr)
        
        num = InStr(1, printArea, ":")
        subStr = Left(printArea, num - 1)
        printArea = Mid(printArea, num + 2)
        startRowPrint = CLng(subStr)
        
        num = InStr(1, printArea, "$")
        subStr = Left(printArea, num - 1)
        printArea = Mid(printArea, num + 1)
        endColPrint = StringColumnToLong(subStr)
        
        endRowPrint = CLng(printArea)
        If startCol < startColPrint Then startCol = startColPrint
        If endCol > endColPrint Then endCol = endColPrint
    Else
        startRowPrint = endRow
        endRowPrint = startRow
    End If
    
    For num = startRow To startRowPrint - 1
        If IsBlankRow(curSheet, num, startCol, endCol) Then
            startRow = num + 1
        Else
            Exit For
        End If
    Next
    For num = endRow To endRowPrint + 1 Step -1
        If IsBlankRow(curSheet, num, startCol, endCol) Then
            endRow = num - 1
        Else
            Exit For
        End If
    Next
    
End Sub

Public Sub GetExcelLastCol(curSheet As Object, ByRef startCol As Long, ByRef endCol As Long)
    Dim endRow As Long
    Dim iCol As Long
    Dim excelRng As Object
    
    Set excelRng = curSheet.Cells.SpecialCells(xlCellTypeLastCell)
    
    startCol = 1
    endRow = excelRng.Row
    endCol = excelRng.Column + excelRng.Count - 1
    
    If endCol > 1 Then
        For iCol = 1 To 20
            If Not IsBlankColumn(curSheet, iCol, endRow) Then
                startCol = iCol
                Exit For
            End If
        Next
        If startCol > 0 Then
            If endCol > constMaxCol Then
                endCol = constMaxCol
                For iCol = startCol + 1 To endCol
                    If IsBlankColumn(curSheet, iCol, endRow) Then
                        endCol = iCol - 1
                        Exit For
                    End If
                Next
            Else
                For iCol = endCol To startCol + 1 Step -1
                    If IsBlankColumn(curSheet, iCol, endRow) Then
                        endCol = iCol - 1
                    Else
                        Exit For
                    End If
                Next
            End If
        Else
            endCol = 0
        End If
    ElseIf endCol = 1 Then
        If IsBlankColumn(curSheet, 1, endRow) Then endCol = 0
    End If
    
End Sub

Public Function IsBlankColumn(curSheet As Object, colNum As Long, endRow As Long) As Boolean
    Dim excelRng As Object
    Dim iRow As Long
    Dim areaRng As Object
    Dim lastRow As Long
    
    On Error Resume Next
    Set excelRng = Nothing
    Set excelRng = curSheet.Range(curSheet.Cells(1, colNum), curSheet.Cells(endRow, colNum)).SpecialCells(xlCellTypeFormulas)
    If Not excelRng Is Nothing Then Exit Function
    
    Set excelRng = Nothing
    Set excelRng = curSheet.Range(curSheet.Cells(1, colNum), curSheet.Cells(endRow, colNum)).SpecialCells(xlCellTypeConstants)
    If excelRng Is Nothing Then
        IsBlankColumn = True
        Exit Function
    End If
    
    For Each areaRng In excelRng.Areas
        lastRow = areaRng.Row + areaRng.Rows.Count - 1
        For iRow = areaRng.Row To lastRow
            If Trim(curSheet.Cells(iRow, colNum).text) <> "" Then Exit Function
        Next
    Next
    
    IsBlankColumn = True
End Function

Public Function IsBlankRow(curSheet As Object, rowNum As Long, startCol As Long, endCol As Long) As Boolean
    Dim excelRng As Object
    Dim iCol As Long
    Dim areaRng As Object
    Dim lastCol As Long
    
    On Error Resume Next
    Set excelRng = Nothing
    Set excelRng = curSheet.Range(curSheet.Cells(rowNum, startCol), curSheet.Cells(rowNum, endCol)).SpecialCells(xlCellTypeFormulas)
    If Not excelRng Is Nothing Then Exit Function
    
    Set excelRng = Nothing
    Set excelRng = curSheet.Range(curSheet.Cells(rowNum, startCol), curSheet.Cells(rowNum, endCol)).SpecialCells(xlCellTypeConstants)
    If excelRng Is Nothing Then
        IsBlankRow = True
        Exit Function
    End If
    
    For Each areaRng In excelRng.Areas
        lastCol = areaRng.Column + areaRng.Columns.Count - 1
        For iCol = areaRng.Column To lastCol
            If Trim(curSheet.Cells(rowNum, iCol).text) <> "" Then Exit Function
        Next
    Next
    
    IsBlankRow = True
End Function

Public Function StringColumnToLong(columnName As String) As Long
    Dim subStr1 As String, subStr2 As String
    Dim colNum As Long
    
    colNum = 0
    subStr1 = UCase(columnName)
    Do While subStr1 <> ""
        subStr2 = Left(subStr1, 1)
        subStr1 = Mid(subStr1, 2)
        
        colNum = colNum * 26 + Asc(subStr2) - 64
    Loop
    
    StringColumnToLong = colNum
End Function

Public Function ContextIsRun() As Boolean
    Dim WinTitle As String * 256, cnt  As Long, hwnd As Long
    Const GW_HWNDNEXT = 2
    Const GW_CHILD = 5

    hwnd = GetDesktopWindow&
    hwnd = GetWindow(hwnd, GW_CHILD)
    Do While hwnd <> 0
        cnt = GetWindowText(hwnd, WinTitle, 255)
        If InStr(1, WinTitle, "Context", vbTextCompare) = 1 Then
            ContextIsRun = True
            Exit Function
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
End Function
