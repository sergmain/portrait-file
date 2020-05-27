Attribute VB_Name = "modSort"
Option Explicit

Public Sub AddSortCollection(col As Collection, s As String)
    Dim i As Long
    Dim res As Long
    Dim name As String
    
    name = GetName(s)
    
    For i = 1 To m_colLogDouble.Count
        If GetName(m_colLogDouble.Item(i)) = name Then
            Call m_colLogDouble.Add(s)
            Exit Sub
        End If
    Next
    
    For i = 1 To col.Count
        res = StringCompare(name, GetName(col(i)))
        If res < 0 Then
            Call col.Add(s, , i)
            Exit Sub
        End If
        If res = 0 Then
            Call m_colLogDouble.Add(col.Item(i))
            Call col.Remove(i)
            Call m_colLogDouble.Add(s)
            Exit Sub
        End If
    Next
    Call col.Add(s)
End Sub

Private Function StringCompare(s1 As String, s2 As String) As Long
    Dim sub1 As String, sub2 As String
    Dim n1 As Long, n2 As Long
    Dim n11 As Long, n21 As Long
    Dim sN1 As String, sN2 As String
    Dim sub11 As String, sub21 As String
    Dim val1 As Long, val2 As Long
    Dim f1 As Boolean, f2 As Boolean
    
    
    n1 = 1: n2 = 1
    Do While n1 <= Len(s1) And n2 <= Len(s2)
        sub1 = Mid(s1, n1, 1)
        sub2 = Mid(s2, n2, 1)
        If sub1 >= "0" And sub1 <= "9" Then
            If sub2 >= "0" And sub2 <= "9" Then
                sN1 = sub1
                sN2 = sub2
                n1 = n1 + 1: n2 = n2 + 1
                f1 = True: f2 = True
                Do While f1 Or f2
                    If f1 Then
                        If n1 <= Len(s1) Then
                            sub1 = Mid(s1, n1, 1)
                            If sub1 >= "0" And sub1 <= "9" Then
                                sN1 = sN1 + sub1
                                n1 = n1 + 1
                            Else
                                f1 = False
                            End If
                        Else
                            f1 = False
                        End If
                    End If
                    If f2 Then
                        If n2 <= Len(s2) Then
                            sub2 = Mid(s2, n2, 1)
                            If sub2 >= "0" And sub2 <= "9" Then
                                sN2 = sN2 + sub2
                                n2 = n2 + 1
                            Else
                                f2 = False
                            End If
                        Else
                            f2 = False
                        End If
                    End If
                    If n1 > Len(s1) And n2 > Len(s2) Then Exit Do
                Loop
                val1 = Val(sN1)
                val2 = Val(sN2)
                If val1 < val2 Then
                    StringCompare = -1
                    Exit Function
                ElseIf val1 > val2 Then
                    StringCompare = 1
                    Exit Function
                End If
            Else
                StringCompare = -1
                Exit Function
            End If
        Else
            If sub2 >= "0" And sub2 <= "9" Then
                StringCompare = 1
                Exit Function
            Else
                If sub1 < sub2 Then
                    StringCompare = -1
                    Exit Function
                ElseIf sub1 > sub2 Then
                    StringCompare = 1
                    Exit Function
                Else
                    n1 = n1 + 1
                    n2 = n2 + 1
                End If
            End If
        End If
    Loop
    
    If n1 > Len(s1) Then
        If n2 > Len(s2) Then
            StringCompare = 0
        Else
            StringCompare = -1
        End If
    Else
        StringCompare = 1
    End If
End Function

