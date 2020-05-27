Attribute VB_Name = "modUtils"
Option Explicit

Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long

Public colMonthes As Collection
Public colMonthesShort As New Collection
Public colMonthesBad As Collection

Public colDocumentTypes1 As New Collection
Public colNotPril As New Collection

Public Sub LoadSettings()
    Dim numFileHandle
    Dim readStr As String
    Dim num As Long, num2 As Long

    numFileHandle = FreeFile
    On Error GoTo err
    Open CheckPath(App.path) + GetName(App.EXEName) + ".ini" For Input As #numFileHandle
    Do While Not EOF(numFileHandle)
        Line Input #numFileHandle, readStr
        num = InStr(1, readStr, "//")
        num2 = InStr(1, readStr, ";")
        If num2 > 0 And (num2 < num Or num = 0) Then num = num2

        If num > 0 Then
            readStr = Trim(Left(readStr, num - 1))
        End If
        
        If VBA.InStr(1, readStr, "Input=", vbTextCompare) = 1 Then
            m_sInputPath = CheckPath(Mid(readStr, Len("Input=") + 1))
        ElseIf VBA.InStr(1, readStr, "Rar=", vbTextCompare) = 1 Then
            m_sRarPath = CheckPath(Mid(readStr, Len("Rar=") + 1))
        ElseIf VBA.InStr(1, readStr, "MaxLen=", vbTextCompare) = 1 Then
            m_maxLen = StringToLong(Mid(readStr, Len("MaxLen=") + 1)) * 1024 * 1024
        ElseIf VBA.InStr(1, readStr, "ExcelLimits=", vbTextCompare) = 1 Then
            maxExcelRow = StringToLong(Mid(readStr, Len("ExcelLimits=") + 1))
        ElseIf VBA.InStr(1, readStr, "LogLargeExcel=", vbTextCompare) = 1 Then
            m_sLogLargeExcelPath = CheckPath(Mid(readStr, Len("LogLargeExcel=") + 1))
        End If
    Loop

err:
    On Error Resume Next
    Close #numFileHandle
End Sub


Public Function CheckPath(path As String)
    If Right(path, 1) <> "\" Then path = path + "\"
    CheckPath = path
End Function


Public Function GetParamArray() As String()
    Dim arrayStr() As String, cnt As Long
    Dim comStr As String
    ReDim arrayStr(0 To 0) As String
    Dim num As Long
    
    comStr = TrimAll(Command())
    
    Do While comStr <> ""
        If Left(comStr, 1) = """" Then
            comStr = Mid(comStr, 2)
            num = InStr(1, comStr, """")
        Else
            num = InStr(1, comStr, " ")
        End If
        
        cnt = cnt + 1
        ReDim Preserve arrayStr(0 To cnt) As String
        If num > 0 Then
            arrayStr(cnt) = Left(comStr, num - 1)
            comStr = TrimAll(Mid(comStr, num + 1))
        Else
            arrayStr(cnt) = comStr
            comStr = ""
        End If
    Loop
    
    GetParamArray = arrayStr
End Function

Public Function GetName(fileName As String) As String
    Dim iExt As Integer
    
    iExt = VBA.InStrRev(fileName, ".")
    If iExt > 0 Then
        GetName = VBA.Left(fileName, iExt - 1)
    Else
        GetName = fileName
    End If
End Function

Public Function GetExtention(fileName As String) As String
    Dim iExt As Integer
    
    iExt = VBA.InStrRev(fileName, ".")
    If iExt = 0 Then
        GetExtention = "."
    Else
        GetExtention = VBA.Mid(fileName, iExt)
    End If
End Function

Public Function GetNameAndExtention(fileName As String) As String
    Dim iExt As Integer
    iExt = VBA.InStrRev(fileName, "\")
    If iExt <> 0 Then
        GetNameAndExtention = VBA.Mid(fileName, iExt + 1)
    Else
        GetNameAndExtention = fileName
    End If
End Function

Public Function GetPath(fileName As String) As String
    Dim iExt As Integer
    iExt = VBA.InStrRev(fileName, "\")
    If iExt <> 0 Then
        GetPath = VBA.Left(fileName, iExt - 1)
    Else
        GetPath = ""
    End If
End Function


Public Function StringToLong(inpstr As String, Optional checkSpace As Boolean = True, Optional defaultValue As Long = 0) As Long
    Dim tStr As String
    Dim num As Long
    tStr = TrimAll(inpstr)
    If checkSpace Then
        num = InStr(1, tStr, " ", vbTextCompare)
        If num > 0 Then tStr = Left(tStr, num - 1)
    End If
    On Error Resume Next
    StringToLong = defaultValue
    StringToLong = CLng(tStr)
End Function

Public Function StringToDouble(inpstr As String, Optional checkSpace As Boolean = True, Optional defaultValue As Long = 0) As Double
    Dim tStr As String
    Dim num As Long
    tStr = TrimAll(inpstr)
    If checkSpace Then
        num = InStr(1, tStr, " ", vbTextCompare)
        If num > 0 Then tStr = Left(tStr, num - 1)
    End If
    On Error Resume Next
    StringToDouble = defaultValue
    StringToDouble = CDbl(tStr)
    If StringToDouble = 0 Then StringToDouble = CDbl(Replace(tStr, ".", ","))
End Function

Public Function TrimAll(ByVal str As String) As String
    Dim i As Long
    str = Replace(str, Chr(7), " ")
    str = Replace(str, Chr(160), " ")
    str = Replace(str, Chr(11), " ")
    str = Replace(str, Chr(1), "")
    str = Trim(str)
    i = InStr(1, str, Chr(13), vbBinaryCompare)
    Do While i = 1
        str = Trim(Mid(str, 2))
        i = InStr(1, str, Chr(13), vbBinaryCompare)
    Loop

    i = InStrRev(str, Chr(13), -1, vbBinaryCompare)
    Do While i > 0 And i = Len(str)
        str = Trim(Left(str, i - 1))
        i = InStrRev(str, Chr(13), -1, vbBinaryCompare)
    Loop

    TrimAll = str
End Function

Public Sub RecProcessFiles(inputPath As String, rec As Boolean, arrExt() As String, colFiles As Collection)
    Dim iExt As Long, iDir As Long
    Dim fileName As String, ext As String
    Dim colDir As New Collection
    
    On Error Resume Next
    m_frmProg.lblProgress = "Поиск файлов"
    m_frmProg.lblStep.Caption = "Найдено: 0"
    Call m_frmProg.Refresh
    
    Dim fso As Object
    Dim fFolder As Object, fFile As Object
    Dim extention As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    m_frmProg.lblProgress = "Поиск файлов"
    m_frmProg.lblStep.Caption = "Чтение каталога"
    Call m_frmProg.Refresh
    Set fFolder = fso.GetFolder(FolderPath:=inputPath)
    m_frmProg.lblStep.Caption = "Найдено: 0"
    Call m_frmProg.Refresh
    For Each fFile In fFolder.Files
        If (colFiles.Count Mod 250) = 0 Then
            m_frmProg.lblProgress = "Поиск файлов"
            m_frmProg.lblStep.Caption = "Найдено: " + CStr(colFiles.Count)
            Call m_frmProg.Refresh
        End If
        fileName = fFile.name
        extention = GetExtention(fileName)
        For iExt = 1 To UBound(arrExt)
            If InStr(1, extention, arrExt(iExt), vbTextCompare) = 2 Then
                If arrExt(iExt) = "xml" Then
                    If IsXml2003(inputPath + fileName) Then
                        Call m_colLogNormal.Add(fileName)
                        Call m_colOutFiles.Add(fileName)
                    Else
                        Call AddSortCollection(colFiles, fileName)
                    End If
                Else
                    Call AddSortCollection(colFiles, fileName)
                End If
                extention = ""
                Exit For
            End If
        Next
        If extention <> "" Then Call m_colLogIngnore.Add(fileName)
    Next fFile
End Sub

Public Sub RecProcessFiles2(inputPath As String, rec As Boolean, arrExt() As String, colFiles As Collection)
    Dim iExt As Long, iDir As Long
    Dim fileName As String, ext As String
    Dim colDir As New Collection
    
    On Error Resume Next
    m_frmProg.lblProgress = "Поиск файлов"
    m_frmProg.lblStep.Caption = "Найдено: 0"
    Call m_frmProg.Refresh
    
    For iExt = 1 To UBound(arrExt)
        If (colFiles.Count Mod 250) = 0 Then
            m_frmProg.lblProgress = "Поиск файлов"
            m_frmProg.lblStep.Caption = "Найдено: " + CStr(colFiles.Count)
            Call m_frmProg.Refresh
        End If
        
        fileName = VBA.Dir(inputPath & "*." + arrExt(iExt))
        Do While fileName <> ""
                ext = ""
                ext = m_colWorkList(fileName)
                If ext = "" Then
                    Call AddSortCollection(colFiles, fileName)
                End If
            fileName = VBA.Dir
        Loop
    Next

    If Not rec Then Exit Sub

    fileName = VBA.Dir(inputPath & "*.*", vbDirectory)
    Do While fileName <> ""
        If GetAttr(inputPath & fileName) And vbDirectory _
            And fileName <> "." And fileName <> ".." Then Call colDir.Add(CheckPath(inputPath & fileName))
        fileName = VBA.Dir
    Loop

    For iDir = 1 To colDir.Count
        Call RecProcessFiles(colDir.Item(iDir), rec, arrExt, colFiles)
    Next
End Sub

Public Function GetWordVersion() As Long
    Dim sVersion As String, dVersion As Long
    Dim num As Long
    
    sVersion = m_WordApp.Version
    num = InStr(1, sVersion, ".")
    If num > 0 Then sVersion = Left(sVersion, num - 1)
    GetWordVersion = StringToLong(sVersion)
End Function



Public Function FileIsPril(doc As Object) As Boolean
    Dim rngFind As Object
    Dim strRange As String, text As String
    Dim iArray  As Long
    Dim cntPar As Long
    Dim curTables As Object
    Dim notPril As String
    
    Set curTables = doc.Paragraphs.First.Range.Tables
    If curTables.Count > 0 Then
        Set rngFind = curTables(1).Range
    Else
        cntPar = 20
        If cntPar > doc.Paragraphs.Count Then cntPar = doc.Paragraphs.Count
        
        Set rngFind = doc.Range(0, doc.Paragraphs(cntPar).Range.End)
        If rngFind.Tables.Count > 0 Then
            cntPar = 100
            If cntPar > doc.Paragraphs.Count Then cntPar = doc.Paragraphs.Count
            Set rngFind = doc.Range(0, doc.Paragraphs(cntPar).Range.End)
        End If
    End If
    strRange = rngFind.text
    strRange = Replace(strRange, Chr(9), " ")
    strRange = TrimAll(strRange)
    For iArray = 1 To UBound(m_arrayPrils)
        text = m_arrayPrils(iArray)
        If UCase(text) = UCase(Left(strRange, Len(text))) Then
            FileIsPril = True
            Exit Function
        End If
    Next
    
    For iArray = 1 To UBound(m_arrayPrilsAdd)
        text = m_arrayPrilsAdd(iArray)
        If UCase(text) = UCase(Left(strRange, Len(text))) Then
            FileIsPril = True
            Exit Function
        End If
    Next
    
    On Error Resume Next
    For iArray = 1 To colDocumentTypes1.Count
        text = UCase(colDocumentTypes1(iArray))
        If text = UCase(Left(strRange, Len(text))) Then
            notPril = ""
            notPril = colNotPril(text)
            If notPril = "" Then
                FileIsPril = True
                Exit Function
            End If
        End If
    Next
End Function

Public Sub ToDoubleCollection(colFiles As Collection)
    If colFiles.Count = 0 Then Exit Sub
    Dim fullFileName As String
    Dim fileName As String, path As String
    Dim iFile As Long
    path = m_sInputPath + "\DOUBLE"
    If Dir(path, vbDirectory) = "" Then MkDir (path)
    
    On Error Resume Next
    For iFile = 1 To colFiles.Count
        fullFileName = m_sInputPath + colFiles(iFile)
        fileName = colFiles(iFile) 
        
        Name fullFileName As path + "\" + fileName
        If Dir(fullFileName) Then Kill (fullFileName)
    Next
End Sub


Public Sub ToBAKCollection(colFiles As Collection)
    Dim fullFileName As String
    Dim fileName As String, path As String
    Dim iFile As Long
    path = m_sInputPath + "\BAK"
    If Dir(path, vbDirectory) = "" Then MkDir (path)
    
    On Error Resume Next
    For iFile = 1 To colFiles.Count
        fullFileName = m_sInputPath + colFiles(iFile)
        fileName = colFiles(iFile)
        If LCase(GetExtention(fileName)) <> "xml" Then
            Name fullFileName As path + "\" + fileName
            If Dir(fullFileName) Then Kill (fullFileName)
        End If
    Next
End Sub

Public Sub ToBAKFile(fileName As String)
    If LCase(GetExtention(fileName)) = ".xml" Then Exit Sub

    Dim fullFileName As String, path As String
    Dim iFile As Long
    path = m_sInputPath + "BAK"
    If Dir(path, vbDirectory) = "" Then MkDir (path)
    
    On Error Resume Next
    fullFileName = m_sInputPath + fileName
    
    Name fullFileName As path + "\" + fileName
    If Dir(fullFileName) Then Call Kill(fullFileName)
End Sub

Public Sub AddMainDoc(doc As Object, srcName As String)
    Dim cnt As Long, i As Long
    Dim strRange As String, tStr As String, sub2 As String
    Dim pars As Object, par As Object
    cnt = 25
    Set pars = doc.Paragraphs
    If pars.Count < cnt Then cnt = pars.Count
    Set par = pars.First
    For i = 1 To cnt
        tStr = par.Range.text
        If Left(tStr, 1) = """" Or Left(tStr, 1) = Chr(171) Then tStr = Mid(tStr, 2)
        If UCase(Left(tStr, 2)) = "О " Then Exit For
        If UCase(Left(tStr, 3)) = "ОБ " Then Exit For
        strRange = strRange + tStr + vbCrLf
        Set par = par.Next
    Next
    
    Dim curMainRecv As New clsMainRecv
    curMainRecv.sOrgan = Replace(strRange, " ", "")
    curMainRecv.m_sFileNameSrc = srcName
    curMainRecv.m_sFileNameDest = doc.name
    Call m_colGlobalMains.Add(curMainRecv)
    Call doc.Close(False)
End Sub


Public Function AddExcel1(fileName As String) As Boolean
    Dim tDoc As Object, errMsg As String
    Dim bHide As Boolean
    Set tDoc = GetExcelDocContent1(fileName, errMsg, bHide)
    If tDoc Is Nothing Or errMsg <> "" Then
        AddExcel1 = False
        If errMsg <> "" Then
            Call AddColByKey(m_colLogBad, GetNameAndExtention(fileName) + " - " + errMsg, GetNameAndExtention(fileName))
        Else
            Call AddColByKey(m_colLogBad, GetNameAndExtention(fileName), GetNameAndExtention(fileName))
        End If
        Exit Function
    End If


    Dim table As Object
    Set table = tDoc.Range.Tables(1)
    Dim strRange As String, strRow As String
    Dim iRow As Long, iCol As Long
    
    
    strRange = ""
    Dim maxRow As Long
    maxRow = 5
    If maxRow > table.Rows.Count Then maxRow = table.Rows.Count
    
    strRow = ""
    On Error Resume Next
    For iRow = 1 To maxRow
        For iCol = 1 To table.Columns.Count
            strRow = strRow + " " + table.Cell(iRow, iCol).Range.text
        Next
    Next
    
    Dim outFileName As String
    outFileName = GetName(fileName) + ".xml"
    Call tDoc.SaveAs(fileName:=outFileName, FileFormat:=11, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False)
    Call tDoc.Close
    strRow = TrimAll(strRow)
    
    Dim curRecv As clsPrilRecv
    Set curRecv = GetRecv(strRow, False)
    If curRecv.m_sDate = "" Or curRecv.m_sNumber = "" Then
        Call m_colLogPril.Add(GetNameAndExtention(fileName) + "(" + GetNameAndExtention(outFileName) + ")")
        AddExcel1 = True
        Exit Function
    End If
    
    curRecv.m_sFileNameSrc = GetNameAndExtention(fileName)
    curRecv.m_sFileNameDest = GetNameAndExtention(outFileName)
    curRecv.m_bExcel = True
    curRecv.m_bHide = bHide
    Call m_colGlobalPrils.Add(curRecv)
    AddExcel1 = True
End Function

Public Sub AddWord(doc As Object, srcName As String, colPrils As Collection, force As Boolean)
    Dim rng As Object
    Set rng = TrimRangeLeft(doc.Range)
    
    Dim iPar As Long
    Dim parRange As Object
    Dim pars As Object
    Dim strRange As String, strRow As String
    strRange = ""
    Set pars = rng.Paragraphs
    For iPar = 1 To 25
        Set parRange = pars.Item(iPar).Range
        strRow = TrimAll(parRange.text)
        If strRow = "" Then
            If parRange.Tables.Count = 0 Then Exit For
        Else
            strRange = strRange + strRow + " "
        End If
    Next
    
    Dim curRecv As clsPrilRecv
    Set curRecv = GetRecv(strRange, True)
    
    If Not force And (curRecv.m_sDate = "" Or curRecv.m_sNumber = "") Then
        Call m_colLogPril.Add(srcName + "(" + GetNameAndExtention(doc.FullName) + ")")
        Exit Sub
    End If
    curRecv.m_sFileNameSrc = srcName
    curRecv.m_sFileNameDest = GetNameAndExtention(doc.FullName)
    If force Then
        Call AddNewPrilToMain(colPrils, curRecv)
    Else
        Call colPrils.Add(curRecv)
    End If
    Call doc.Close(False)
    
End Sub

Public Function GetRecv(strRange As String, onlyPrilNumber As Boolean) As clsPrilRecv
    Dim numK As Long, numDate As Long, numNumber As Long
    Dim numPril As Double
    Dim newRecv As clsPrilRecv
    Set newRecv = New clsPrilRecv
    strRange = Replace(strRange, Chr(9), " ")
    strRange = Replace(strRange, Chr(10), " ")
    strRange = replaceAll(strRange, "  ", " ")
    strRange = TrimAll(strRange)
    
    If InStr(1, strRange, "Приложение", vbTextCompare) = 1 Then
        Dim subStr As String
        subStr = Trim(Mid(strRange, 11))
        If Left(subStr, 1) = "№" Then
            subStr = Trim(Mid(subStr, 2))
        ElseIf Left(subStr, 1) = "N" Then
            subStr = Trim(Mid(subStr, 2))
        End If
        
        numPril = StringToDouble(subStr)
        numK = InStr(1, strRange, " к ", vbTextCompare)
        If numK = 0 Then
            numK = InStr(1, strRange, " ", vbTextCompare)
            strRange = Mid(strRange, numK + 1)
        Else
            strRange = Mid(strRange, numK + 3)
        End If
    Else
        numPril = 0
        numK = InStr(1, strRange, " ", vbTextCompare)
        strRange = Mid(strRange, numK + 1)
    End If
    
    numDate = InStr(1, strRange, " от ", vbTextCompare)
    numNumber = InStr(1, strRange, "№", vbTextCompare)
    If numNumber = 0 Then numNumber = InStr(1, strRange, "N", vbTextCompare)
    
    newRecv.m_PrilNumber = numPril
    If onlyPrilNumber Then
        Set GetRecv = newRecv
        Exit Function
    End If
    If numDate = 0 And numNumber = 0 Then
        If onlyPrilNumber Then Set GetRecv = newRecv
        Exit Function
    End If
    
    Dim strS As String, strOrgan As String
    Dim strNumber As String, strDate As String
    If numDate > 0 Then
        If numNumber = 0 Then
            strDate = Mid(strRange, numDate)
        ElseIf numNumber < numDate Then
            strS = Mid(strRange, numNumber)
            strNumber = Mid(strRange, numNumber, numDate - numNumber)
            strDate = Mid(strRange, numDate)
            strOrgan = Left(strRange, numNumber - 1)
        Else
            strS = Mid(strRange, numDate)
            strDate = Mid(strRange, numDate, numNumber - numDate)
            strNumber = Mid(strRange, numNumber)
            strOrgan = Left(strRange, numDate - 1)
        End If
    Else
        If numNumber > 0 Then
            strS = Mid(strRange, numNumber)
            strNumber = Mid(strRange, numNumber)
            strOrgan = Left(strRange, numNumber - 1)
        Else
            strOrgan = strRange
        End If
    End If
    
    strDate = replaceAll(Trim(strDate), "  ", " ")
    strNumber = replaceAll(Trim(strNumber), "  ", " ")
    strS = replaceAll(Trim(strS), "  ", " ")
    
    Call GetSubDate(strS, strDate, strNumber)
    Dim strDateD As String
    If strDate <> "" Then strDateD = ConvertDateToDDDD(strDate)
    
    newRecv.m_sDate = Replace(strDate, " ", "")
    newRecv.m_sDateD = strDateD
    newRecv.m_sNumber = Replace(strNumber, " ", "")
    strOrgan = Trim(strOrgan)
    strOrgan = Replace(strOrgan, "  ", " ")
    newRecv.m_sOrgan = strOrgan
    
    Set GetRecv = newRecv
    
End Function

Public Function ConvertDateToDDDD(strDate As String) As String
    Dim num As Long, iMon As Long
    Dim newDate As String
    Dim mon As String
    num = InStr(1, strDate, " ")
    newDate = Left(strDate, num - 1) + "."
    For iMon = 1 To 12
        If InStr(1, strDate, colMonthes(iMon), vbTextCompare) > 0 Then
            newDate = newDate + colMonthesShort(iMon) + "."
            Exit For
        End If
    Next
    num = InStr(num + 1, strDate, " ")
    newDate = newDate + Mid(strDate, num + 1)
    ConvertDateToDDDD = newDate
End Function

Public Sub AddNewPrilToMain(colPrils As Collection, newPril As clsPrilRecv)
    Dim iPril As Long
    Dim curPril As clsPrilRecv
    
    For iPril = 1 To colPrils.Count
        Set curPril = colPrils(iPril)
        If newPril.m_PrilNumber < curPril.m_PrilNumber Then
            Call colPrils.Add(newPril, , iPril)
            Exit Sub
        End If
    Next
    Call colPrils.Add(newPril)
End Sub


Public Function FileIsExcel2010(fileName As String) As Integer
    Dim numFileHandle
    Dim buf(1 To 4096) As Byte
    Dim find As String
    Dim iArr As Long, cntArr As Long, iString As Long, cntString As Long
    
    FileIsExcel2010 = 0
    
    numFileHandle = FreeFile
    Open fileName For Binary Access Read As #numFileHandle
    Get #numFileHandle, , buf
    Close #numFileHandle
    
    find = "workbook"
    cntString = Len(find)
    cntArr = UBound(buf) - cntString
    
    For iArr = 1 To cntArr
        FileIsExcel2010 = 1
        For iString = 1 To Len(find)
            If buf(iArr + iString - 1) <> Asc(Mid(find, iString, 1)) Then
                FileIsExcel2010 = 0
                Exit For
            Else
                iArr = iArr
            End If
        Next
        If FileIsExcel2010 > 0 Then
            If buf(15) = 223 Then FileIsExcel2010 = 2
            Exit Function
        End If
    Next
    
End Function

Public Function FileIsExcel2003(fileName As String) As Boolean
    Dim numFileHandle
    Dim buf(1 To 513) As Byte
    Dim iArr As Long, cntArr As Long, iString As Long, cntString As Long
    
    numFileHandle = FreeFile
    Open fileName For Binary Access Read As #numFileHandle
    Get #numFileHandle, , buf
    Close #numFileHandle
    
    If buf(65) = 0 Then
        FileIsExcel2003 = True
    ElseIf buf(65) = 2 Then
        FileIsExcel2003 = True
    Else
        FileIsExcel2003 = False
    End If
End Function


Public Sub AddToCollection(col As Collection, addStr As String)
    Dim iCol As Long
    Dim withoutExtencion As String
    
    withoutExtencion = GetName(addStr)
    
    For iCol = 1 To col.Count
        If withoutExtencion < GetName(col(iCol)) Then
            Call col.Add(addStr, , iCol)
            Exit Sub
        End If
    Next
    Call col.Add(addStr)
End Sub


Public Function TrimRangeLeft(rng As Object, Optional checkTable As Boolean = True) As Object
    Dim i As Long, startRange As Long, endRange As Long
    Dim str As String
    Dim parentDoc As Object
    Set parentDoc = rng.Parent
    
    str = rng.text

    startRange = rng.start: endRange = rng.End
    Dim trimSymbols(1 To 6) As String, iTrim As Long
    trimSymbols(1) = " ": trimSymbols(2) = Chr(13): trimSymbols(3) = Chr(7)
    trimSymbols(4) = Chr(12): trimSymbols(5) = Chr(1): trimSymbols(6) = Chr(11)

    On Error GoTo ErrLabel
    If (Not checkTable) Or parentDoc.Range(startRange, startRange + 1).Tables.Count = 0 Then
        For iTrim = 1 To UBound(trimSymbols)
            i = InStr(1, str, trimSymbols(iTrim), vbBinaryCompare)
            If i = 1 Then Exit For
        Next
        Do While i = 1
            startRange = startRange + 1
            If checkTable And parentDoc.Range(startRange, startRange + 1).Tables.Count > 0 Then Exit Do
            str = parentDoc.Range(startRange, endRange).text
            For iTrim = 1 To UBound(trimSymbols)
                i = InStr(1, str, trimSymbols(iTrim), vbBinaryCompare)
                If i = 1 Then Exit For
            Next
        Loop
    End If

    Set TrimRangeLeft = parentDoc.Range(startRange, endRange)
    Exit Function
ErrLabel:
    If startRange = endRange Then
        If endRange < parentDoc.Range.End Then
            endRange = endRange + 1
        Else
            startRange = startRange - 1
        End If
    End If
    On Error Resume Next
    Set TrimRangeLeft = parentDoc.Range(startRange, endRange)
End Function


Private Function GetFirstNotBlankCellString(curTable As Object) As String
    Dim curCell As Object
    Dim strCell As String
    Dim rng As Object
    Dim par As Object
    Dim pars As Object
    Dim num As Long
    
    Set curCell = curTable.Cell(1, 1)
    Do While Not curCell Is Nothing
        strCell = TrimAll(curCell.Range)
        If strCell <> "" Then
            strCell = TrimAll(curCell.Range)
            Set rng = curCell.Range
            Set pars = rng.Paragraphs
            Set par = pars.First
            Do While Not par Is Nothing
                strCell = TrimAllWithOut11(par.Range.text)
                
                If strCell <> "" Then
                    GetFirstNotBlankCellString = strCell
                    Exit Function
                End If
                Set par = par.Next
            Loop
            
            GetFirstNotBlankCellString = TrimAllWithOut11(curCell.Range)
            Exit Function
        End If
        Set curCell = curCell.Next
    Loop
    GetFirstNotBlankCellString = ""
End Function

Public Function replaceAll(ByVal str As String, findText As String, replacedText As String) As String
    Dim lenStr As Long
    
    Do
        lenStr = Len(str)
        str = Replace(str, findText, replacedText)
    Loop While lenStr <> Len(str)
    replaceAll = str
End Function



Public Sub GetSubDate(ByVal str As String, ByRef documentDate, ByRef documentNumber)
    Dim strRange As String
    Dim startDate As Long, iChar As Long, iDate As Long
    Dim strEndRange As String
    
    strRange = str
    strEndRange = ""
    strRange = Replace(strRange, "_", "")
    If InStr(1, strRange, "от ", vbTextCompare) Then
        strRange = Replace(strRange, "«", " ")
    Else
        strRange = Replace(strRange, "«", "от ")
    End If
    strRange = Replace(strRange, "»", " ")
    strRange = Replace(strRange, """", " ")
    strRange = Replace(strRange, Chr(147), " ")
    strRange = Replace(strRange, Chr(148), " ")
    strRange = Replace(strRange, "  ", " ")
    strRange = Replace(strRange, "  ", " ")
    startDate = InStr(1, strRange, "N")
    If startDate = 0 Then startDate = InStr(1, strRange, "№")
    If startDate > 0 Then
        documentNumber = TrimAll(Replace(Mid(strRange, startDate + 1), Chr(7), ""))
        strRange = Trim(Left(strRange, startDate - 1))
        
        If GetDatePostion(strRange) = 0 Then
            startDate = InStr(1, documentNumber, "от ", vbTextCompare)
            If startDate > 0 Then
                strRange = Mid(documentNumber, startDate)
                documentNumber = TrimAll(Left(documentNumber, startDate - 1))
            End If
        End If
        startDate = InStr(1, documentNumber, Chr(13))
        If startDate > 0 Then
            strEndRange = TrimAll(Mid(documentNumber, startDate + 1))
            documentNumber = TrimAll(Left(documentNumber, startDate - 1))
        End If
    Else
        startDate = InStr(1, documentNumber, "от ", vbTextCompare)
        If startDate <> 1 Then
            startDate = InStr(1, strRange, "года", vbTextCompare)
            If startDate = 0 Then startDate = InStr(1, strRange, "год", vbTextCompare)
            If startDate > 0 Then
                startDate = startDate + 3
                documentNumber = TrimAll(Mid(strRange, startDate + 1))
            Else
                documentNumber = ""
            End If
            
        Else
            documentNumber = TrimAll(Left(TrimAll(strRange), startDate - 1))
        End If
    End If
        
    startDate = InStr(1, strRange, "от ", vbTextCompare)
    If startDate > 0 Then
        strRange = Mid(strRange, startDate)
    End If
    
    If GetDatePostion(strRange) = 0 And GetDatePostion(strEndRange) > 0 Then strRange = strEndRange
    startDate = InStr(1, strRange, "год", vbTextCompare)
    If startDate = 0 Then startDate = InStr(1, strRange, "г.", vbTextCompare)
    If startDate > 0 Then strRange = Trim(Left(strRange, startDate - 1))

    startDate = Len(strRange)
    For iChar = "0" To "9"
        iDate = InStr(1, strRange, iChar)
        If iDate > 0 And iDate < startDate Then startDate = iDate
    Next
    If startDate <> Len(strRange) Then
        strRange = Mid(strRange, startDate)
    Else
        strRange = ""
    End If
    
    If (Not IsMonthExist(strRange)) And (Len(Replace(strRange, " ", "")) = 8 Or Len(Replace(strRange, " ", "")) = 9) Then
        strRange = Replace(strRange, " ", "")
        strRange = Replace(strRange, ".", "")
        strRange = Left(strRange, 2) + "." + Mid(strRange, 3, 2) + "." + Mid(strRange, 5)
    End If
    
    documentDate = ConvertDate(strRange)
    documentNumber = replaceAll(documentNumber, "  ", " ")
    
    iChar = InStr(1, documentNumber, "Об ", vbTextCompare)
    If iChar = 0 Then iChar = InStr(1, documentNumber, "О ", vbTextCompare)
    If iChar > 0 Then documentNumber = Trim(Left(documentNumber, iChar - 1))
    
End Sub

Private Function ConvertDate(ByVal strDate As String) As String
    Dim numDot As Long
    Dim numDate As Long
    
    numDot = InStr(1, strDate, ".")
    If numDot > 0 Then
        strDate = Replace(strDate, " ", "")
        numDot = InStr(1, strDate, ".")
        ConvertDate = Left(strDate, numDot - 1) + " "
        strDate = Mid(strDate, numDot + 1)
        numDot = InStr(1, strDate, ".")
        If numDot > 0 Then
            On Error Resume Next
            numDate = 0
            numDate = CLng(Left(strDate, numDot - 1))
            On Error GoTo 0
            ConvertDate = ConvertDate + GetMonth(numDate) + " "
            On Error Resume Next
            numDate = 0
            numDate = CLng(Mid(strDate, numDot + 1, 4))
            If numDate = 0 Then numDate = CLng(Mid(strDate, numDot + 1, 2))
            On Error GoTo 0
            If numDate = 0 Then
                ConvertDate = ConvertDate + "год!!!"
            Else
                If numDate < 2000 Then numDate = numDate + 2000
                ConvertDate = ConvertDate + CStr(numDate)
            End If
        Else
            ConvertDate = ConvertDate + "месяц!!! год!!!"
        End If
    Else
        ConvertDate = Trim(strDate)
        numDate = GetDatePostion(ConvertDate)
        If numDate > 0 Then numDate = InStr(numDate + 4, ConvertDate, " ")
        If numDate > 0 Then ConvertDate = Left(ConvertDate, numDate - 1)
        For numDate = 1 To colMonthes.Count
            numDot = InStr(1, ConvertDate, colMonthes(numDate), vbTextCompare)
            If numDot > 1 Then
                If Mid(ConvertDate, numDot - 1, 1) <> "" Then ConvertDate = Left(ConvertDate, numDot - 1) + " " + Mid(ConvertDate, numDot)
                Exit For
            End If
        Next
        For numDate = 1 To colMonthes.Count
            numDot = InStr(1, ConvertDate, colMonthesBad(numDate), vbTextCompare)
            If numDot > 1 Then
                ConvertDate = Left(ConvertDate, numDot - 1) + colMonthes(numDate) + Mid(ConvertDate, Len(colMonthesBad(numDate)) + numDot)
                Exit For
            End If
        Next
        
    End If

    ConvertDate = TrimAll(ConvertDate)
    If Left(ConvertDate, 1) = "0" Then ConvertDate = Mid(ConvertDate, 2)

    ConvertDate = replaceAll(ConvertDate, "  ", " ")
    numDot = InStrRev(ConvertDate, " ")
    If numDot > 0 Then
        If (Len(ConvertDate) - numDot) = 5 Then
            If Mid(ConvertDate, numDot + 2, 1) = "0" Then
                ConvertDate = Left(ConvertDate, numDot + 1) + Mid(ConvertDate, numDot + 3)
            End If
        End If
    End If
End Function


Private Function TrimAllWithOut11(ByVal str As String) As String
    Dim i As Long
    str = Replace(str, Chr(7), " ")
    str = Replace(str, Chr(160), " ")
    str = Replace(str, Chr(1), "")
    str = Trim(str)
    i = InStr(1, str, Chr(13), vbBinaryCompare)
    Do While i = 1
        str = Trim(Mid(str, 2))
        i = InStr(1, str, Chr(13), vbBinaryCompare)
    Loop

    i = InStrRev(str, Chr(13), -1, vbBinaryCompare)
    Do While i > 0 And i = Len(str)
        str = Trim(Left(str, i - 1))
        i = InStrRev(str, Chr(13), -1, vbBinaryCompare)
    Loop

    TrimAllWithOut11 = str
End Function


Public Function GetDatePostion(ByVal strRange As String) As Long
    Dim iYear As Long, iMonth As Long
    Dim position As Long
    Dim position2 As Long
    Dim trimedStr As String
    
    trimedStr = TrimAll(strRange)
    Do While InStr(1, trimedStr, " ")
        trimedStr = Replace(trimedStr, " ", "")
    Loop
    
    For iYear = 2001 To 2030
        position = InStr(1, trimedStr, CStr(iYear))
        If position = 0 Then
            position = InStr(1, trimedStr, CStr(iYear + 18000))
        End If
        
        If position > 0 Then
            position2 = InStr(1, trimedStr, ".")
            If position2 > 0 And position2 < position Then
                GetDatePostion = position
                Exit Function
            End If
            For iMonth = 1 To 12
                position2 = InStr(1, trimedStr, colMonthesShort(iMonth), vbTextCompare)
                If position2 > 0 And position2 < position Then
                    GetDatePostion = position
                    Exit Function
                End If
            Next
            For iMonth = 1 To 12
                position2 = InStr(1, trimedStr, colMonthes(iMonth), vbTextCompare)
                If position2 > 0 And position2 < position Then
                    GetDatePostion = position
                    Exit Function
                End If
            Next
            For iMonth = 1 To 12
                position2 = InStr(1, trimedStr, colMonthesBad(iMonth), vbTextCompare)
                If position2 > 0 And position2 < position Then
                    GetDatePostion = position
                    Exit Function
                End If
            Next
            Exit For
        End If
    Next
    
    GetDatePostion = 0
End Function

Private Function IsMonthExist(str As String) As Boolean
    If InStr(1, str, "января", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "февраля", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "марта", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "апреля", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "мая", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "июня", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "июля", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "августа", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "сентября", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "октября", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "ноября", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    If InStr(1, str, "декабря", vbTextCompare) > 0 Then IsMonthExist = True: Exit Function
    
    IsMonthExist = False
End Function

Private Function GetMonth(month As Long) As String
    Select Case month
        Case 1
            GetMonth = "января"
        Case 2
            GetMonth = "февраля"
        Case 3
            GetMonth = "марта"
        Case 4
            GetMonth = "апреля"
        Case 5
            GetMonth = "мая"
        Case 6
            GetMonth = "июня"
        Case 7
            GetMonth = "июля"
        Case 8
            GetMonth = "августа"
        Case 9
            GetMonth = "сентября"
        Case 10
            GetMonth = "октября"
        Case 11
            GetMonth = "ноября"
        Case 12
            GetMonth = "декабря"
        Case Else
            GetMonth = "месяц!!!"
    End Select
End Function


Public Function IsXml2003(fileName As String) As Boolean
    Dim numFileHandle
    Dim readStr As String
    On Error GoTo ErrLabel
    numFileHandle = FreeFile
    Open fileName For Input As #numFileHandle
    Input #numFileHandle, readStr
    Input #numFileHandle, readStr
    Input #numFileHandle, readStr
    Close #numFileHandle
    readStr = Left(readStr, 300)
    
    IsXml2003 = InStr(1, readStr, "pkg:package", vbTextCompare) = 0
    Exit Function
ErrLabel:
    On Error Resume Next
    Close #numFileHandle
    IsXml2003 = False
End Function

Public Sub LoadTypesFromCfg()
    Dim readStr As String, path As String
    Dim num As Long, num2 As Long
    Dim numFileHandle
    
    
    Set colDocumentTypes1 = New Collection
    numFileHandle = FreeFile
    On Error GoTo err
    Open CheckPath(App.path) + GetName(App.EXEName) + ".types" For Input As #numFileHandle
    Do While Not EOF(numFileHandle)
        Line Input #numFileHandle, readStr
        readStr = Trim(readStr)
        num = InStr(1, readStr, "//")
        num2 = InStr(1, readStr, ";")
        If num2 > 0 And num2 < num Then num = num2
        If num > 0 Then
            readStr = Trim(Left(readStr, num - 1))
        End If
        If readStr <> "" Then
            Call colDocumentTypes1.Add(LCase(readStr), LCase(readStr))
        End If
    Loop

err:
    On Error Resume Next
    Close #numFileHandle
End Sub


