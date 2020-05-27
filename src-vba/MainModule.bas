Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Const wdOpenFormatAuto = 0

Public m_sInputPath As String
Public m_sRarPath As String

Public m_arrayPrils() As String
Public m_arrayPrilsAdd() As String
Public m_colNotPril As Collection

Public m_colWorkList As New Collection

Public m_maxLen As Long
Public m_frmProg As frmProgress

Public m_bHideLists As Boolean
Public m_bNoCollate As Boolean
Public m_WordApp As Object

Public m_colGlobalMains As New Collection
Public m_colGlobalPrils As New Collection

Public m_colLogNormal As New Collection
Public m_colLogPril As New Collection
Public m_colLogBad As New Collection
Public m_colOutFiles As New Collection
Public m_colLogDouble As New Collection
Public m_colLogIngnore As New Collection

Private Const maxRead = 100000
Public maxExcelRow As Long
Public m_sLogLargeExcelPath As String


Public Sub Main()
    Dim params() As String
    
    maxExcelRow = 0
    
    Call LoadSettings
    Set colMonthes = New Collection
    Call colMonthes.Add("января"): Call colMonthes.Add("февраля"): Call colMonthes.Add("марта")
    Call colMonthes.Add("апреля"): Call colMonthes.Add("мая"): Call colMonthes.Add("июня")
    Call colMonthes.Add("июля"): Call colMonthes.Add("августа"): Call colMonthes.Add("сентября")
    Call colMonthes.Add("октября"): Call colMonthes.Add("ноября"): Call colMonthes.Add("декабря")
    
    Set colMonthesShort = New Collection
    Call colMonthesShort.Add("01"): Call colMonthesShort.Add("02"): Call colMonthesShort.Add("03")
    Call colMonthesShort.Add("04"): Call colMonthesShort.Add("05"): Call colMonthesShort.Add("06")
    Call colMonthesShort.Add("07"): Call colMonthesShort.Add("08"): Call colMonthesShort.Add("09")
    Call colMonthesShort.Add("10"): Call colMonthesShort.Add("11"): Call colMonthesShort.Add("12")
    
    Set colMonthesBad = New Collection
    Call colMonthesBad.Add("январь"): Call colMonthesBad.Add("февраль"): Call colMonthesBad.Add("марть")
    Call colMonthesBad.Add("апрель"): Call colMonthesBad.Add("май"): Call colMonthesBad.Add("июнь")
    Call colMonthesBad.Add("июль"): Call colMonthesBad.Add("августь"): Call colMonthesBad.Add("сентябрь")
    Call colMonthesBad.Add("октябрь"): Call colMonthesBad.Add("ноябрь"): Call colMonthesBad.Add("декабрь")
    
    Set colNotPril = New Collection
    Call colNotPril.Add("ПОСТАНОВЛЕНИЕ", "ПОСТАНОВЛЕНИЕ")
    Call colNotPril.Add("РАСПОРЯЖЕНИЕ", "РАСПОРЯЖЕНИЕ")
    Call colNotPril.Add("ЗАКОН", "ЗАКОН")
    Call colNotPril.Add("УКАЗ", "УКАЗ")
    Call colNotPril.Add("ПРИКАЗ", "ПРИКАЗ")
    
    If m_sInputPath = "" Then m_sInputPath = CheckPath(CurDir)
    
    params = GetParamArray
    
    Dim iParam As Long, tParam As Long
    For iParam = 1 To UBound(params)
        If UCase(params(iParam)) = "/HIDE" Then
            m_bHideLists = True
        ElseIf UCase(params(iParam)) = "/NOCOLLATE" Then
            m_bNoCollate = True
        ElseIf UCase(params(iParam)) = "/NOEXCELLIMITS" Then
            maxExcelRow = 0
        ElseIf UCase(Left(params(iParam), Len("/EXCELLIMITS="))) = "/EXCELLIMITS=" Then
            tParam = StringToLong(Mid(params(iParam), Len("/EXCELLIMITS=") + 1), False, -1)
            If tParam >= 0 Then maxExcelRow = tParam
        Else
            m_sInputPath = CheckPath(params(iParam))
        End If
    Next
    
    Call ProcessFiles
    
    Call ExitProcess(0)
End Sub



Private Sub ProcessFiles()
    
    Dim arrExt(1 To 6) As String
    arrExt(1) = "doc": arrExt(2) = "rtf": arrExt(3) = "xml": arrExt(4) = "dot": arrExt(5) = "xls": arrExt(6) = "xlt"

    ReDim m_arrayPrils(1 To 10) As String
    m_arrayPrils(1) = "Приложени"
    m_arrayPrils(2) = "Утвержд"
    m_arrayPrils(3) = "Одобрен"
    m_arrayPrils(4) = "Пояснительн"
    m_arrayPrils(5) = "Рекомендован"
    m_arrayPrils(6) = "Определен"
    m_arrayPrils(7) = "Принято "
    m_arrayPrils(8) = "Принят "
    m_arrayPrils(9) = "Приняты "
    m_arrayPrils(10) = "Установлен"
    
    ReDim m_arrayPrilsAdd(1 To 3) As String
    m_arrayPrilsAdd(1) = "Согласов"
    m_arrayPrilsAdd(2) = "К "
    m_arrayPrilsAdd(3) = "ОПОВЕЩЕНИЕ"
    
    Set m_colNotPril = New Collection
    Call m_colNotPril.Add("ПОСТАНОВЛЕНИЕ", "ПОСТАНОВЛЕНИЕ")
    Call m_colNotPril.Add("РАСПОРЯЖЕНИЕ", "РАСПОРЯЖЕНИЕ")
    Call m_colNotPril.Add("ЗАКОН", "ЗАКОН")
    Call m_colNotPril.Add("УКАЗ", "УКАЗ")
    Call m_colNotPril.Add("ПРИКАЗ", "ПРИКАЗ")
    Dim colFiles As New Collection
    Set m_frmProg = New frmProgress
    Call m_frmProg.Show
    Set m_colLogDouble = New Collection
    Set m_colLogNormal = New Collection
    Set m_colLogPril = New Collection
    Set m_colLogBad = New Collection
    Set m_colGlobalPrils = New Collection
    Set m_colGlobalMains = New Collection
    Set m_colLogIngnore = New Collection
    Set m_colOutFiles = New Collection
    
    Call RecProcessFiles(m_sInputPath, False, arrExt, colFiles)
    Call ToDoubleCollection(m_colLogDouble)
    
    If colFiles.Count = 0 Then Exit Sub
    
    Set m_WordApp = CreateObject("Word.Application")
    m_WordApp.Visible = False
    On Error GoTo LabelErr
    
    Dim m_WdRecoveryType  As Long
    If GetWordVersion() > 12 Then
        m_WdRecoveryType = 19
    Else
        m_WdRecoveryType = 16
    End If

    
    
    Dim firstFile As Boolean, findMain As Boolean, isPril As Boolean
    Dim iFile As Long, iSubFile As Long, iPril As Long, num As Long
    Dim fileName As String, checkName As String, mainFileName As String, extention As String
    Dim colPril As Collection, colDocPrils1 As Collection
    Dim docMain1 As New clsMainDoc
    Dim tStr As String, errMsg As String
    Dim tDoc As Object, rng As Object
    Dim bHide As Boolean
    Dim colToBak As Collection
    Dim curPril As clsPrilRecv
    Dim bInsertMethod As Boolean
    
    
    firstFile = True
    Dim logMain As String
    For iFile = 1 To colFiles.Count
        Call PrintLogs
        m_frmProg.lblProgress.Caption = "Файл " + CStr(iFile) + " из " + CStr(colFiles.Count) + ": " + GetNameAndExtention(colFiles(iFile))
        m_frmProg.lblStep.Caption = ""
        Call m_frmProg.Refresh
        
        fileName = colFiles(iFile)
        tStr = GetName(colFiles(iFile))
        
        Dim checkName2 As String, mainFileName2 As String, findMain2 As Boolean
        Dim checkName3 As String
        Dim lng As Long
        
        Set colPril = New Collection
        num = InStrRev(tStr, "-")
        checkName2 = ""
        If num > 0 Then
            lng = StringToLong(Mid(tStr, num + 1))
            If lng > 0 And lng < 10 Then
                checkName = Left(tStr, num)
                findMain = True
                mainFileName = ""
                
                checkName2 = tStr + "-"
                findMain2 = False
                mainFileName2 = colFiles(iFile)
            Else
                checkName = tStr + "-"
                checkName3 = tStr + "-П"
                findMain = False
                mainFileName = colFiles(iFile)
            End If
        Else
            checkName = tStr + "-"
            checkName3 = tStr + "-П"
            findMain = False
            mainFileName = colFiles(iFile)
        End If
        
        For iSubFile = iFile + 1 To colFiles.Count
            If UCase(Left(colFiles(iSubFile), Len(checkName))) = UCase(checkName) Then
                tStr = Mid(colFiles(iSubFile), Len(checkName) + 1)
                lng = StringToLong(GetName(Mid(colFiles(iSubFile), Len(checkName) + 1)))
                If lng > 0 And lng < 100 Then
                    Call colPril.Add(colFiles(iSubFile))
                    m_frmProg.lblProgress.Caption = m_frmProg.lblProgress.Caption + " +" + colFiles(iSubFile)
                    Call m_frmProg.Refresh
                    iFile = iFile + 1
                Else
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        
        If iSubFile = iFile + 1 And checkName2 <> "" Then
            checkName = checkName2
            findMain = findMain2
            mainFileName = mainFileName2
            For iSubFile = iFile + 1 To colFiles.Count
                If UCase(Left(colFiles(iSubFile), Len(checkName))) = UCase(checkName) Then
                    tStr = Mid(colFiles(iSubFile), Len(checkName) + 1)
                    lng = StringToLong(GetName(Mid(colFiles(iSubFile), Len(checkName) + 1)))
                    If lng > 0 And lng < 100 Then
                        Call colPril.Add(colFiles(iSubFile))
                        m_frmProg.lblProgress.Caption = m_frmProg.lblProgress.Caption + " +" + colFiles(iSubFile)
                        Call m_frmProg.Refresh
                        iFile = iFile + 1
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If iSubFile = iFile + 1 And checkName3 <> "" Then
            checkName2 = checkName
            checkName = checkName3
            For iSubFile = iFile + 1 To colFiles.Count
                If UCase(Left(colFiles(iSubFile), Len(checkName))) = UCase(checkName) Then
                    tStr = Mid(colFiles(iSubFile), Len(checkName) + 1)
                    lng = StringToLong(GetName(Mid(colFiles(iSubFile), Len(checkName) + 1)))
                    If lng > 0 And lng < 100 Then
                        Call colPril.Add(colFiles(iSubFile))
                        m_frmProg.lblProgress.Caption = m_frmProg.lblProgress.Caption + " +" + colFiles(iSubFile)
                        Call m_frmProg.Refresh
                        iFile = iFile + 1
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            checkName = checkName2
        End If
        
        If findMain Then
            Call colPril.Add(fileName)
        End If
        
            
        Set colDocPrils1 = New Collection
        If colPril.Count > 0 Then
            Set colToBak = New Collection
            bInsertMethod = False
            If mainFileName <> "" Then
                Set tDoc = m_WordApp.Documents.Add
                tDoc.ActiveWindow.Visible = False
                errMsg = MyInsertFile1(tDoc, m_sInputPath + mainFileName)
                If errMsg = "" Then
                    Set docMain1 = New clsMainDoc
                    Set docMain1.m_doc = tDoc
                    docMain1.m_sFileNameSrc = GetNameAndExtention(mainFileName)
                    Call colToBak.Add(mainFileName)
                Else
                    Call m_colLogBad.Add(mainFileName + " - " + errMsg)
                    Call tDoc.Close(False)
                End If
            End If
        
            DoEvents
            For iPril = 1 To colPril.Count
                m_frmProg.lblStep.Caption = "Чтение файлов"
                m_frmProg.Refresh
                extention = GetExtention(colPril(iPril))
                If InStr(1, extention, ".xl", vbTextCompare) <> 1 Then
                    Set tDoc = m_WordApp.Documents.Add
                    tDoc.ActiveWindow.Visible = False
                    errMsg = MyInsertFile1(tDoc, m_sInputPath + colPril(iPril))
                    If errMsg = "" Then
                        If findMain Then
                            isPril = FileIsPril(tDoc)
                        Else
                            isPril = True
                        End If
                        If Not isPril Then
                            mainFileName = colPril(iPril)
                            Set docMain1.m_doc = tDoc
                            docMain1.m_sFileNameSrc = mainFileName
                            findMain = False
                        Else
                            Call tDoc.SaveAs(fileName:=m_sInputPath + GetName(colPril(iPril)) + ".xml", FileFormat:=11, _
                                LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                                :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                                SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                                False)
                            Call AddWord(tDoc, GetNameAndExtention(colPril(iPril)), colDocPrils1, True)
                            bInsertMethod = True
                        End If
                        Call colToBak.Add(colPril(iPril))
                    Else
                        Call m_colLogBad.Add(colPril(iPril) + " - " + errMsg)
                        Call tDoc.Close(False)
                    End If
                Else
                    errMsg = ExcelToDoc(m_sInputPath + colPril(iPril), bHide, colDocPrils1, True)
                    If errMsg <> "" Then
                        If errMsg <> "" Then
                            Call AddColByKey(m_colLogBad, colPril(iPril) + " - " + errMsg, colPril(iPril))
                        Else
                            Call AddColByKey(m_colLogBad, colPril(iPril), colPril(iPril))
                        End If
                    Else
                        Call colToBak.Add(colPril(iPril))
                    End If
                End If
            Next
            
            If Right(checkName, 1) = "-" Then checkName = Left(checkName, Len(checkName) - 1)
            checkName = checkName + ".xml"
            If mainFileName <> "" Then
                m_frmProg.lblStep.Caption = "Добавление приложений в главный документ"
                m_frmProg.Refresh
                If bInsertMethod Then
                    Set tDoc = docMain1.m_doc
                    logMain = docMain1.m_sFileNameSrc + "(" + GetNameAndExtention(checkName) + ")"
                    For iPril = 1 To colDocPrils1.Count
                        Set curPril = colDocPrils1(iPril)
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Set rng = tDoc.Range(tDoc.Range.End - 1, tDoc.Range.End)
                        Call rng.InsertBreak(2)
                        Set rng = tDoc.Range(tDoc.Range.End - 1, tDoc.Range.End)
                        rng.InsertFile fileName:=m_sInputPath + curPril.m_sFileNameDest, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
                        logMain = logMain + " + " + curPril.m_sFileNameSrc
                        If curPril.m_bHide Then logMain = logMain + "[скрытые листы]"
                    Next
                    m_frmProg.lblStep.Caption = "Сохранение файла"
                    m_frmProg.Refresh
                    Call tDoc.SaveAs(fileName:=m_sInputPath + checkName, FileFormat:=11, _
                        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                        False)
                    Call tDoc.Close(False)
                    For iPril = 1 To colDocPrils1.Count
                        Set curPril = colDocPrils1(iPril)
                        Call Kill(m_sInputPath + curPril.m_sFileNameDest)
                    Next
                Else
                    m_frmProg.lblStep.Caption = "Сохранение файла"
                    m_frmProg.Refresh
                    Call tDoc.SaveAs(fileName:=m_sInputPath + checkName, FileFormat:=11, _
                        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                        False)
                    Call tDoc.Close(False)
                    logMain = docMain1.m_sFileNameSrc + "(" + GetNameAndExtention(checkName) + ")" + SpliceFiles(m_sInputPath + checkName, colDocPrils1)
                End If
                Call colPril.Add(fileName)
                Call ToBAKCollection(colToBak)
                Call m_colLogNormal.Add(logMain)
                Call m_colOutFiles.Add(checkName)
            Else
                m_frmProg.lblStep.Caption = "Склейка приложений в один документ"
                m_frmProg.Refresh
                Set tDoc = m_WordApp.Documents.Add
                tDoc.ActiveWindow.Visible = False
                If bInsertMethod Then
                    logMain = ""
                    For iPril = 1 To colDocPrils1.Count
                        Set curPril = colDocPrils1(iPril)
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Call tDoc.Range.InsertParagraphAfter
                        Set rng = tDoc.Range(tDoc.Range.End - 1, tDoc.Range.End)
                        Call rng.InsertBreak(2)
                        Set rng = tDoc.Range(tDoc.Range.End - 1, tDoc.Range.End)
                        rng.InsertFile fileName:=m_sInputPath + curPril.m_sFileNameDest, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
                        logMain = logMain + " + " + curPril.m_sFileNameSrc
                        If curPril.m_bHide Then logMain = logMain + "[скрытые листы]"
                    Next
                    logMain = checkName + " : " + Mid(logMain, 4)
                    m_frmProg.lblStep.Caption = "Сохранение файла"
                    m_frmProg.Refresh
                    Call tDoc.SaveAs(fileName:=m_sInputPath + checkName, FileFormat:=11, _
                        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                        False)
                    Call tDoc.Close(False)
                    For iPril = 1 To colDocPrils1.Count
                        Set curPril = colDocPrils1(iPril)
                        Call Kill(m_sInputPath + curPril.m_sFileNameDest)
                    Next
                    Call m_colLogPril.Add(logMain)
                    Call m_colOutFiles.Add(checkName)
                Else
                    m_frmProg.lblStep.Caption = "Сохранение файла"
                    m_frmProg.Refresh
                    
                    Call tDoc.SaveAs(fileName:=m_sInputPath + checkName, FileFormat:=11, _
                        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                        False)
                    Call tDoc.Close(False)
                    logMain = checkName + " : " + Mid(SpliceFiles(m_sInputPath + checkName, colDocPrils1), 4)
                    Call m_colLogPril.Add(logMain)
                    Call m_colOutFiles.Add(checkName)
                End If
                Call ToBAKCollection(colToBak)
            End If
        Else
            m_frmProg.lblStep.Caption = "Чтение файла"
            m_frmProg.Refresh
            extention = GetExtention(fileName)
            If InStr(1, extention, ".xl", vbTextCompare) <> 1 Then
                Set tDoc = m_WordApp.Documents.Add
                tDoc.ActiveWindow.Visible = False
                errMsg = MyInsertFile1(tDoc, m_sInputPath + fileName)
                If errMsg = "" Then
                    m_frmProg.lblStep.Caption = "Сохранение файла"
                    m_frmProg.Refresh
                    Call tDoc.SaveAs(fileName:=m_sInputPath + GetName(fileName) + ".xml", FileFormat:=11, _
                        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
                        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                        False)
                    If Not m_bNoCollate Then
                        m_frmProg.lblStep.Caption = "Определение реквизитов"
                        m_frmProg.Refresh
                        If Not FileIsPril(tDoc) Then
                            Call AddMainDoc(tDoc, GetNameAndExtention(fileName))
                        Else
                            Call AddWord(tDoc, GetNameAndExtention(fileName), m_colGlobalPrils, False)
                        End If
                    Else
                        Call m_colLogNormal.Add(GetNameAndExtention(fileName) + "(" + GetName(fileName) + ".xml)")
                        Call tDoc.Close(False)
                        Call m_colOutFiles.Add(GetName(fileName) + ".xml")
                    End If
                    Call ToBAKFile(fileName)
                Else
                    Call m_colLogBad.Add(mainFileName + " - " + errMsg)
                    Call tDoc.Close(False)
                End If
            Else
                If Not m_bNoCollate Then
                    errMsg = ExcelToDoc(m_sInputPath + fileName, bHide, m_colGlobalPrils, False)
                Else
                    errMsg = ExcelToDoc(m_sInputPath + fileName, bHide, colDocPrils1, True)
                End If
                If errMsg <> "" Then
                    If errMsg <> "" Then
                        Call AddColByKey(m_colLogBad, GetNameAndExtention(fileName) + " - " + errMsg, GetNameAndExtention(fileName))
                    Else
                        Call AddColByKey(m_colLogBad, GetNameAndExtention(fileName), GetNameAndExtention(fileName))
                    End If
                Else
                    If m_bNoCollate Then
                        If bHide Then
                            Call m_colLogPril.Add(GetNameAndExtention(fileName) + "(" + GetName(fileName) + ".xml)[скрытые листы]")
                        Else
                            Call m_colLogPril.Add(GetNameAndExtention(fileName) + "(" + GetName(fileName) + ".xml)")
                        End If
                        Call m_colOutFiles.Add(GetName(fileName) + ".xml")
                    End If
                    Call ToBAKFile(fileName)
                End If
            End If
        End If
    Next
    
    Call PrintLogs
LabQuit:
    On Error Resume Next
    Call m_WordApp.Quit(False)
    If Not m_oExcel Is Nothing Then
        Call m_oExcel.Quit
        Set m_oExcel = Nothing
    End If
    
    m_frmProg.Hide
    Set m_frmProg = Nothing


    Dim lRar As Long
    Dim sessionDate As Date
    sessionDate = Now()
    lRar = ExecRar(m_colOutFiles, sessionDate)
    Exit Sub
    
LabelErr:
    On Error Resume Next
    Call m_WordApp.Quit(False)
    If Not m_oExcel Is Nothing Then
        Call m_oExcelMacro.Close(False)
        Call m_oExcel.Quit
    End If
    
    m_frmProg.Hide
    Set m_frmProg = Nothing

End Sub

Public Sub PrintLogs()
    Dim numFileHandle
    Dim iCol As Long
    
    numFileHandle = FreeFile
    Open m_sInputPath + "PortraitFile.log" For Output As #numFileHandle

    If m_colLogNormal.Count > 0 Then
        Print #numFileHandle, "====== Преобразованные файлы"
        For iCol = 1 To m_colLogNormal.Count
            Print #numFileHandle, m_colLogNormal(iCol)
        Next
        Print #numFileHandle, ""
    End If
    
    If m_colLogPril.Count > 0 Then
        Print #numFileHandle, "====== Файлы приложений, для которых не нашлось главного файла"
        For iCol = 1 To m_colLogPril.Count
            Print #numFileHandle, m_colLogPril(iCol)
        Next
        Print #numFileHandle, ""
    End If
    
    If m_colLogBad.Count > 0 Then
        Print #numFileHandle, "====== Файлы, которые не удалось преобразовать"
        For iCol = 1 To m_colLogBad.Count
            Print #numFileHandle, m_colLogBad(iCol)
        Next
        Print #numFileHandle, ""
    End If
    
    If m_colLogDouble.Count > 0 Then
        Print #numFileHandle, "====== Файлы с дублируемыми именами (перенесены в папку DOUBLE)"
        For iCol = 1 To m_colLogDouble.Count
            Print #numFileHandle, m_colLogDouble(iCol)
        Next
        Print #numFileHandle, ""
    End If
    
    If m_colLogIngnore.Count > 0 Then
        Print #numFileHandle, "====== Файлы, обработка которых не производилась"
        For iCol = 1 To m_colLogIngnore.Count
            Print #numFileHandle, m_colLogIngnore(iCol)
        Next
        Print #numFileHandle, ""
    End If
    
    
    Close #numFileHandle

End Sub


Public Sub AddColByKey(col As Collection, value As String, key As String)
    On Error Resume Next
    Call col.Add(value, key)
End Sub


Private Function ExecRar(colFiles As Collection, sessionDate As Date) As Long
    Dim numFileHandle
    Dim iCol As Long
    
    numFileHandle = FreeFile
    Open m_sInputPath + "fileslist" For Output As #numFileHandle
    For iCol = 1 To colFiles.Count
        Print #numFileHandle, colFiles(iCol)
    Next
    Close #numFileHandle
    
    
    Dim tStr As String
    tStr = Format(sessionDate, "yymmdd_hhMMss")
    
    Dim dirPath As String
    If Mid(m_sInputPath, 2, 1) = ":" Then
        Call ChDrive(Mid(m_sInputPath, 1, 2))
    End If
    Call ChDir(m_sInputPath)
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    Call WshShell.Run("""" + m_sRarPath + "winrar.exe"" a -afzip " + tStr + " @fileslist", 1, True)
    
    Call Kill(m_sInputPath + "fileslist")
    iCol = 0
    iCol = FileLen(tStr + ".zip")
    
    ExecRar = iCol
End Function


Private Function MyInsertFile1(ByRef tDoc As Object, ByVal filePath As String) As String
    Dim errMsg As String
    Dim tDoc2 As Object
    On Error Resume Next
    tDoc.Range.InsertFile fileName:=filePath, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
    errMsg = err.Description
    MyInsertFile1 = errMsg
End Function

Private Function SpliceFiles(mainFile As String, colDocPrils1 As Collection) As String
    If colDocPrils1.Count = 0 Then Exit Function
    On Error GoTo ErrLabel
    Dim numFileNumberMain, numFileNumberPril
    Dim mainSize As Long, prilSize As Long, readLen As Long
    Dim bufferString As String
    Dim startBody As Boolean, num As Long, iCol As Long
    Dim inputPath As String
    Dim curPril As clsPrilRecv

    inputPath = CheckPath(GetPath(mainFile))
    numFileNumberMain = FreeFile
    Open mainFile For Binary As numFileNumberMain
    mainSize = LOF(numFileNumberMain)
    Seek numFileNumberMain, mainSize - 25
    
    For iCol = 1 To colDocPrils1.Count
        numFileNumberPril = FreeFile
        Set curPril = colDocPrils1.Item(iCol)
        Open inputPath + curPril.m_sFileNameDest For Binary As numFileNumberPril
        prilSize = LOF(numFileNumberPril) - 26
        readLen = 0
        startBody = False
    
        Do While readLen <= prilSize
            If (readLen + maxRead) > prilSize Then
                bufferString = String(prilSize - readLen, " ")
                readLen = prilSize + 1
            Else
                bufferString = String(maxRead, " ")
                readLen = readLen + maxRead
            End If
            Get numFileNumberPril, , bufferString
            If Not startBody Then
                num = InStr(1, bufferString, "<w:body>")
                If num > 0 Then
                    bufferString = Mid(bufferString, num + 8)
                    startBody = True
                End If
            End If
            If startBody Then
                Put numFileNumberMain, , bufferString
            End If
        Loop
        Close numFileNumberPril
        SpliceFiles = SpliceFiles + " + " + curPril.m_sFileNameSrc
        If curPril.m_bHide Then SpliceFiles = SpliceFiles + "[скрытые листы]"
    Next
    
    Put numFileNumberMain, , "</w:body></w:wordDocument>"
    
    Close numFileNumberMain
    On Error Resume Next
    For Each curPril In colDocPrils1
        Call Kill(inputPath + curPril.m_sFileNameDest)
    Next
    
    Exit Function
ErrLabel:
    On Error Resume Next
    Close numFileNumberMain

    
End Function

