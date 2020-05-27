Attribute VB_Name = "modExcel2"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const constMaxCol = 50
Private Const xlCellTypeLastCell = 11
Private Const xlCellTypeFormulas = -4123
Private Const xlCellTypeConstants = 2
Private Const xlSheetVisible = -1
Private Const xlA1 = 1

Private Const xlEdgeTop = 8
Private Const xlEdgeBottom = 9
Private Const xlEdgeLeft = 7
Private Const xlEdgeRight = 10
Public startTime As Long, endTime As Long


Public Function ExcelToDoc(oFileName As String, ByRef hideListsExist As Boolean, colDocPrils As Collection, onlyPrilNumber As Boolean) As String
    Dim strVersion As String
    On Error GoTo ErrLabel
    If Not m_oExcel Is Nothing Then
        strVersion = m_oExcel.Version
        If strVersion = "" Then
            Set m_oExcel = Nothing
        End If
    End If
    
    Dim dispAlerts As Boolean
    Dim referenceStyle As Long
    Dim macroName As String
    macroName = GetName(App.EXEName) + ".xls"

    If m_oExcel Is Nothing Then
        m_frmProg.lblStep.Caption = "Запуск Excel"
        m_frmProg.Refresh
        Set m_oExcel = CreateObject("Excel.Application")
        If m_oExcel Is Nothing Then
            ExcelToDoc = "Ошибка запуска Excel"
            Exit Function
        End If
        
        m_oExcel.Visible = False
        Set m_oExcelMacro = m_oExcel.WorkBooks.Open(CheckPath(App.path) + macroName)
    End If
    
    
    dispAlerts = m_oExcel.DisplayAlerts
    m_oExcel.DisplayAlerts = False
    hideListsExist = False
    referenceStyle = m_oExcel.referenceStyle
    m_oExcel.referenceStyle = xlA1
    
    Dim excelWorkBook As Object, excelSheets As Object, excelSheet As Object
    
    Dim result As String, recvStr As String, res2 As String
    m_frmProg.lblStep.Caption = "Преобразование в XML: " + GetNameAndExtention(oFileName)
    m_frmProg.Refresh
    result = m_oExcel.Run(macroName + "!ExcelToDoc", oFileName, m_bHideLists, hideListsExist, recvStr, m_frmProg, maxExcelRow)
    If result <> "" Then
        ExcelToDoc = "Ошибка при преобразовании файла (" + result + ")"
        If m_sLogLargeExcelPath <> "" Then
            If InStr(1, result, "Превышен", vbTextCompare) > 0 And InStr(1, result, "лимит", vbTextCompare) > 0 Then
                Call WriteLogLargeExcel(oFileName)
            End If
        End If
        GoTo ErrLabel
    End If
    
    Dim prilRecv  As clsPrilRecv
    Set prilRecv = GetRecv(recvStr, onlyPrilNumber)
    prilRecv.m_bExcel = True
    prilRecv.m_sFileNameSrc = GetNameAndExtention(oFileName)
    prilRecv.m_sFileNameDest = GetName(prilRecv.m_sFileNameSrc) + ".xml"
    prilRecv.m_bHide = hideListsExist
    Call AddNewPrilToMain(colDocPrils, prilRecv)
    
    
    m_oExcel.DisplayAlerts = dispAlerts
    m_oExcel.referenceStyle = referenceStyle
    Exit Function
    
ErrLabel:
    m_oExcel.referenceStyle = referenceStyle
    If ExcelToDoc = "" Then ExcelToDoc = err.Description
    On Error Resume Next
    If Not excelWorkBook Is Nothing Then Call excelWorkBook.Close(False)
    
End Function

Private Sub WriteLogLargeExcel(oFileName As String)
    On Error Resume Next
    Dim numFileHandle
    numFileHandle = FreeFile
    Open m_sLogLargeExcelPath + GetNameAndExtention(oFileName) + ".null" For Output As #numFileHandle
    Close #numFileHandle
End Sub

Public Sub PrintSheet(ByRef sheet As Object, ByVal startRow As Long, ByVal startCol As Long, ByVal endRow As Long, ByVal endCol As Long, ByVal fileName, ByRef numFileHandle, ByRef prilRecv As clsPrilRecv, onlyPrilNumber As Boolean)
    If numFileHandle = -1 Then
        numFileHandle = FreeFile
        Open fileName For Output As #numFileHandle
        Print #numFileHandle, "<?xml version=""1.0"" encoding=""windows-1251"" standalone=""yes""?>"
        Print #numFileHandle, "<?mso-application progid=""Word.Document""?>"
        Print #numFileHandle, "<w:wordDocument xmlns:w=""http://schemas.microsoft.com/office/word/2003/wordml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:sl=""http://schemas.microsoft.com/schemaLibrary/2003/core"" xmlns:aml=""http://schemas.microsoft.com/aml/2001/core"" xmlns:wx=""http://schemas.microsoft.com/office/word/2003/auxHint"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:dt=""uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"" w:macrosPresent=""no"" w:embeddedObjPresent=""no"" w:ocxPresent=""no"" xml:space=""preserve"">"
        Print #numFileHandle, "<w:body>"
    End If
    Print #numFileHandle, "<wx:sect><w:sectPr><w:pgSz w:w=""31185"" w:h=""16840""/><w:pgMar w:top=""567"" w:right=""0"" w:bottom=""567"" w:left=""0"" w:header=""0"" w:footer=""0"" w:gutter=""0""/></w:sectPr><w:tbl>"
    Print #numFileHandle, "<w:tblPr>"
    Print #numFileHandle, "<w:tblW w:type=""auto"" w:w=""0""/>"
    Print #numFileHandle, "<w:tblLayout w:type=""Fixed""/>"
    Print #numFileHandle, "<w:tblInd w:type=""dxa"" w:w=""0""/>"
    Print #numFileHandle, "<w:jc w:val=""left""/>"
    Print #numFileHandle, "<w:tblBorders>"
    Print #numFileHandle, "<w:left w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "<w:top w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "<w:right w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "<w:bottom w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "<w:insideH w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "<w:insideV w:color=""000000"" w:sz=""4"" w:val=""single""/>"
    Print #numFileHandle, "</w:tblBorders>"
    
    Print #numFileHandle, "</w:tblPr>"
    Print #numFileHandle, "<w:tblGrid>"
    
    Dim i As Long, j As Long, iW As Long, curWidth As Long
    Dim curCell As Object, merged As Object, mergeRange As Object
    Dim printCell As Boolean
    Dim mergeProp As String, strPar As String
    Dim cellWidthes() As Long
    
    ReDim cellWidthes(startCol To endCol) As Long
    Dim sumLen As Long, koefWidth As Double
    For j = startCol To endCol
        cellWidthes(j) = sheet.Cells(1, j).ColumnWidth * 0.198 * 567
        sumLen = sumLen + cellWidthes(j)
    Next
    If sumLen > 30000 Then
        koefWidth = CDbl(30000) / CDbl(sumLen)
        For j = startCol To endCol
            cellWidthes(j) = cellWidthes(j) * koefWidth
        Next
    End If
    For j = startCol To endCol
        Print #numFileHandle, "<w:gridCol w:w=""" + CStr(cellWidthes(j)) + """/>"
    Next
    Print #numFileHandle, "</w:tblGrid>"
    
    Dim strBorders As String, strOrientation As String, strAlign As String
    Dim curRow As Long, curCol As Long
    
    startTime = GetTickCount()
    For i = startRow To endRow
        If (i Mod 1000) = 0 Then
            m_frmProg.lblStep.Caption = "Преобразование в XML(" + fileName + ") - " + CStr(i) + " из " + CStr(endRow)
            m_frmProg.Refresh
        End If
        Print #numFileHandle, "<w:tr>"
        For j = startCol To endCol
            Set curCell = sheet.Cells(i, j)
            printCell = True
            mergeProp = ""
            curWidth = cellWidthes(j)
            If curCell.MergeCells Then
                Set merged = curCell.MergeArea
                Set mergeRange = merged.Columns
                If mergeRange.Count > 1 Then
                    If merged.Column = j Then
                        mergeProp = "<w:gridSpan w:val=""" + CStr(mergeRange.Count) + """/>"
                    Else
                        printCell = False
                    End If
                End If
                If printCell Then
                    Set mergeRange = merged.Rows
                    If mergeRange.Count > 1 Then
                        If merged.Row = i Then
                            mergeProp = mergeProp + "<w:vmerge w:val=""restart""/>"
                        Else
                            mergeProp = mergeProp + "<w:vmerge/>"
                        End If
                    End If
                End If
            End If
            If printCell Then
                strBorders = ""
                curRow = curCell.Row: curCol = curCell.Column
                With curCell.Borders
                    If .Item(xlEdgeTop).LineStyle = -4142 Then
                        If curRow > 1 Then
                            If sheet.Cells(curRow - 1, curCol).Borders.Item(xlEdgeBottom).LineStyle = -4142 Then
                                strBorders = strBorders + "<w:top w:val=""nil""/>"
                            End If
                        Else
                            strBorders = strBorders + "<w:top w:val=""nil""/>"
                        End If
                    End If
                    If .Item(xlEdgeLeft).LineStyle = -4142 Then
                        If curCol > 1 Then
                            If sheet.Cells(curRow, curCol - 1).Borders.Item(xlEdgeRight).LineStyle = -4142 Then
                                strBorders = strBorders + "<w:left w:val=""nil""/>"
                            End If
                        Else
                            strBorders = strBorders + "<w:left w:val=""nil""/>"
                        End If
                    End If
                    If .Item(xlEdgeBottom).LineStyle = -4142 Then
                        If sheet.Cells(curRow + 1, curCol).Borders.Item(xlEdgeTop).LineStyle = -4142 Then
                            strBorders = strBorders + "<w:bottom w:val=""nil""/>"
                        End If
                    End If
                    If .Item(xlEdgeRight).LineStyle = -4142 Then
                        If sheet.Cells(curRow, curCol + 1).Borders.Item(xlEdgeLeft).LineStyle = -4142 Then
                            strBorders = strBorders + "<w:right w:val=""nil""/>"
                        End If
                    End If
                End With
                
                If strBorders <> "" Then
                    strBorders = "<w:tcBorders>" + strBorders + "</w:tcBorders>"
                End If
                If curCell.Orientation = -4128 Then
                    strOrientation = ""
                ElseIf curCell.Orientation = -4171 Then
                    strOrientation = "<w:textFlow w:val=""bt-lr""/>"
                ElseIf curCell.Orientation = -4170 Then
                    strOrientation = "<w:textFlow w:val=""tb-rl""/>"
                Else
                    strOrientation = ""
                End If
                
                strPar = GetRanges(curCell)
                Print #numFileHandle, "<w:tc><w:tcPr><w:tcW w:type=""dxa"" w:w=""" + CStr(curWidth) + """/>" + strBorders + mergeProp + strOrientation + "</w:tcPr>"
                If strPar <> "" Then
                        Select Case curCell.HorizontalAlignment
                            Case -4108
                                strAlign = "center"
                            Case -4152
                                strAlign = "right"
                            Case Else
                                strAlign = "left"
                        End Select
                    Print #numFileHandle, "<w:p><w:pPr><w:jc w:val=""" + strAlign + """/></w:pPr>"
                    Print #numFileHandle, strPar
                    Print #numFileHandle, "</w:p>"
                    If prilRecv Is Nothing Then
                        Set prilRecv = GetRecv(curCell.text, onlyPrilNumber)
                    End If
                Else
                    Print #numFileHandle, "<w:p/>"
                End If
                Print #numFileHandle, "</w:tc>"
            End If
        Next
        Print #numFileHandle, "</w:tr>"
    Next
    
    
    Print #numFileHandle, "</w:tbl></wx:sect>"
    
End Sub

Private Function GetRanges(curCell As Object) As String
    Dim cellText As String
    cellText = curCell.text
    If cellText = "" Then Exit Function
    
    Dim curRange As clsRange
    Dim colRanges As New Collection
    Dim i As Long, curLen As Long
    Dim lastFont As Byte, curFont As Byte, lastStart As Long
    Dim num As Long
    Dim addRange As clsRange
    Dim curChars As Object
    
    If curCell.font.Superscript = False Then
        If curCell.font.Subscript = True Then
            Set addRange = New clsRange
            addRange.m_br = False
            addRange.m_font = 2
            addRange.m_text = cellText
            Call colRanges.Add(addRange)
        Else
            Set addRange = New clsRange
            addRange.m_br = False
            addRange.m_font = 0
            addRange.m_text = cellText
            Call colRanges.Add(addRange)
        End If
    ElseIf curCell.font.Superscript = True Then
        Set addRange = New clsRange
        addRange.m_br = False
        addRange.m_font = 1
        addRange.m_text = cellText
        Call colRanges.Add(addRange)
    Else
        curLen = Len(cellText)
        
        Set curChars = curCell.Characters(start:=1, Length:=1)
        If curChars.font.Superscript = True Then
            lastFont = 1
        ElseIf curChars.font.Subscript = True Then
            lastFont = 2
        Else
            lastFont = 0
        End If
        
        
        lastStart = 1
        For i = 2 To curLen
            Set curChars = curCell.Characters(start:=1, Length:=1)
            If curChars.font.Superscript = True Then
                curFont = 1
            ElseIf curChars.font.Subscript = True Then
                curFont = 2
            Else
                curFont = 0
            End If
            If lastFont <> curFont Then
                Set addRange = New clsRange
                addRange.m_br = False
                addRange.m_font = lastFont
                addRange.m_text = Mid(cellText, lastStart, i - lastStart)
                Call colRanges.Add(addRange)
                
                lastFont = curFont
                lastStart = i
            End If
        Next
        Set addRange = New clsRange
        addRange.m_br = False
        addRange.m_font = lastFont
        addRange.m_text = Mid(cellText, lastStart, curLen - lastStart + 1)
        Call colRanges.Add(addRange)
    End If
        
    Dim strLeft As String, strRight As String
    i = 1
    Do While i <= colRanges.Count
        Set curRange = colRanges.Item(i)
        num = InStr(1, curRange.m_text, Chr(10))
        If num > 0 Then
            strLeft = Left(curRange.m_text, num - 1)
            strRight = Mid(curRange.m_text, num + 1)
            
            If strLeft <> "" Then
                curRange.m_text = strLeft
                Set addRange = New clsRange
                addRange.m_br = True
                addRange.m_font = curRange.m_font
                Call colRanges.Add(addRange, , , i)
                If strRight <> "" Then
                    Set addRange = New clsRange
                    addRange.m_text = strRight
                    addRange.m_font = curRange.m_font
                    addRange.m_br = False
                    Call colRanges.Add(addRange, , , i + 1)
                End If
                i = i + 1
            Else
                If strRight <> "" Then
                    Set addRange = New clsRange
                    addRange.m_font = curRange.m_font
                    addRange.m_br = True
                    Call colRanges.Add(addRange, , i)
                    curRange.m_text = strRight
                Else
                    curRange.m_br = True
                End If
            End If
        End If
        i = i + 1
    Loop
    
    If colRanges.Count = 1 Then
        Set curRange = colRanges(1)
        If Trim(curRange.m_text) = "" Then
            Exit Function
        End If
    End If
    Dim strOut As String
    strOut = ""

    lastFont = 3
    For i = 1 To colRanges.Count
        Set curRange = colRanges(i)
        If lastFont <> curRange.m_font Then
            If strOut <> "" Then
                strOut = strOut + "</w:r>"
            End If
            strOut = strOut + "<w:r>"
            If curRange.m_font = 1 Then
                strOut = strOut + "<w:rPr><w:vertAlign w:val=""superscript""/></w:rPr>"
            ElseIf curRange.m_font = 2 Then
                strOut = strOut + "<w:rPr><w:vertAlign w:val=""subscript""/></w:rPr>"
            End If
            lastFont = curRange.m_font
        End If
        If curRange.m_br Then
            strOut = strOut + "<w:br/>"
        Else
            strOut = strOut + "<w:t>" + curRange.m_text + "</w:t>"
        End If
    Next
    If strOut <> "" Then strOut = strOut + "</w:r>"
    GetRanges = strOut
End Function


Private Function GetFont1(curChars As Object) As Long
    If curChars.font.Superscript = True Then
        GetFont1 = 1
    ElseIf curChars.font.Subscript = True Then
        GetFont1 = -1
    Else
        GetFont1 = 0
    End If
End Function
Private Function GetAlign1(curCell As Object) As String
    Select Case curCell.HorizontalAlignment
        Case -4108
            GetAlign1 = "center"
        Case -4152
            GetAlign1 = "right"
        Case Else
            GetAlign1 = "left"
    End Select
End Function

Private Function GetBorders1(curCell As Object, sheet As Object) As String
    Dim curRow As Long, curCol As Long
    curRow = curCell.Row: curCol = curCell.Column
    With curCell.Borders
        If .Item(xlEdgeTop).LineStyle = -4142 Then
            If curRow > 1 Then
                If sheet.Cells(curRow - 1, curCol).Borders.Item(xlEdgeBottom).LineStyle = -4142 Then
                    GetBorders1 = GetBorders1 + "<w:top w:val=""nil""/>"
                End If
            Else
                GetBorders1 = GetBorders1 + "<w:top w:val=""nil""/>"
            End If
        End If
        If .Item(xlEdgeLeft).LineStyle = -4142 Then
            If curCol > 1 Then
                If sheet.Cells(curRow, curCol - 1).Borders.Item(xlEdgeRight).LineStyle = -4142 Then
                    GetBorders1 = GetBorders1 + "<w:left w:val=""nil""/>"
                End If
            Else
                GetBorders1 = GetBorders1 + "<w:left w:val=""nil""/>"
            End If
        End If
        If .Item(xlEdgeBottom).LineStyle = -4142 Then
            If sheet.Cells(curRow + 1, curCol).Borders.Item(xlEdgeTop).LineStyle = -4142 Then
                GetBorders1 = GetBorders1 + "<w:bottom w:val=""nil""/>"
            End If
        End If
        If .Item(xlEdgeRight).LineStyle = -4142 Then
            If sheet.Cells(curRow, curCol + 1).Borders.Item(xlEdgeLeft).LineStyle = -4142 Then
                GetBorders1 = GetBorders1 + "<w:right w:val=""nil""/>"
            End If
        End If
    End With
    
    If GetBorders1 <> "" Then
        GetBorders1 = "<w:tcBorders>" + GetBorders1 + "</w:tcBorders>"
    End If
End Function

Private Function GetOrientaion1(curCell As Object) As String
    If curCell.Orientation = -4171 Then
        GetOrientaion1 = "<w:textFlow w:val=""bt-lr""/>"
    ElseIf curCell.Orientation = -4170 Then
        GetOrientaion1 = "<w:textFlow w:val=""tb-rl""/>"
    End If
End Function

Private Function NewRange1(text As String, font As Long, br As Boolean) As clsRange
    Set NewRange1 = New clsRange
    NewRange1.m_br = br
    NewRange1.m_font = font
    NewRange1.m_text = text
End Function

