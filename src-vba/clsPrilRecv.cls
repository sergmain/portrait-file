VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrilRecv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_sNumber As String
Public m_sDate As String
Public m_sDateD As String
Private sOrgan As String
Private sOrgans() As String
Private iCnt As Long
Public m_PrilNumber As Double

Public m_sFileNameSrc As String
Public m_sFileNameDest As String

Public m_bHide As Boolean

Public m_bExcel As Boolean

Property Let m_sOrgan(s As String)
    sOrgan = s
    sOrgans = Split(s, " ")
    Dim i As Long, clen As Long
    For i = 0 To UBound(sOrgans)
        sOrgans(i) = Replace(sOrgans(i), Chr(171), "")
        sOrgans(i) = Replace(sOrgans(i), Chr(187), "")
        sOrgans(i) = Replace(sOrgans(i), """", "")
        clen = Len(sOrgans(i))
        If clen > 4 Then
            sOrgans(i) = Left(sOrgans(i), clen - 2)
        ElseIf clen > 3 Then
            sOrgans(i) = Left(sOrgans(i), clen - 1)
        End If
    Next
    
    If UBound(sOrgans) < 5 Then
        iCnt = UBound(sOrgans)
    Else
        iCnt = 5
    End If
End Property

Property Get m_sOrgan() As String
    m_sOrgan = sOrgan
End Property

Property Get m_sOrgans() As String()
    m_sOrgans = sOrgans
End Property

Property Get m_iCnt() As Long
    m_iCnt = iCnt
End Property

