Option Explicit
' メンバ変数
Public odd As Boolean
Private data As String
Private length As Long
Private hosei As Long
Private kaisi As Long
Const STS_ODDINI = 11
Const STS_INI = 1
Const STS_SEARCH = 2
Const STS_RBOOT = 3
Const STS_END = 4
Public Function DataInitial(inData As String)
    data = inData
    length = Len(data)
    odd = oddchekc
End Function
Public Function Trans() As String
    Dim status_id As Integer: status_id = STS_INI
    Dim serch As Long: serch = 1
    hosei = 0
    Trans = data
    Do
        If status_id = STS_RBOOT Then status_id = STS_INI '再投入(変換後のズレ対応)
        Trans = TransChk(status_id, serch, Trans)
        If status_id = STS_END And odd Then
            status_id = STS_ODDINI ''再投入(奇数文字のズレ対応)
            odd = False
        End If
    Loop While status_id <> STS_END
End Function
Private Function TransChk(status_id As Integer, serch As Long, Trans As String) As String
    Dim i As Long
    Dim chk1 As String
    Dim chk2 As String
    Dim temp As String
    TransChk = Trans
    For i = serch To length
        temp = Mid(Trans, i, 1)
        Select Case status_id
        Case STS_ODDINI
            If temp = """" Then
                status_id = STS_INI
            End If
        Case STS_INI
            If temp = """" Then
                status_id = STS_SEARCH
                kaisi = i
            End If
        Case STS_SEARCH
            If temp = """" Then
                chk1 = Mid(TransChk, i - 2, 1)
                chk2 = Mid(TransChk, i - 1, 1)
                If chk1 = "A" And chk2 = "M" Then
                    TransChk = TransExe(Trans, i, False)
                    serch = i + hosei + 1
                    hosei = 0
                    status_id = STS_RBOOT
                    i = length
                ElseIf chk1 = "P" And chk2 = "M" Then
                    TransChk = TransExe(Trans, i, True)
                    serch = i + hosei + 1
                    hosei = 0
                    status_id = STS_RBOOT
                    i = length
                Else
                    status_id = STS_INI
                End If
            End If
        End Select
    Next i
    If status_id <> STS_RBOOT Then
        status_id = STS_END
    End If
End Function
'12345678901234567890
'6 13 2021 8:00AM
'2020-6-13 8:00:00
Private Function TransExe(inData As String, i As Long, flg As Boolean) As String
    Dim j As Long
    Dim chk As String
    Dim tmp As Variant
    Dim tmp_time As Variant
    TransExe = Mid(inData, 1, kaisi - 1)
    chk = Mid(inData, kaisi + 1, i - kaisi - 1)
    
    tmp = Split(chk, " ") 'Split("6 13 2021 8:00AM", " ")
    
    chk = """"
    
    chk = chk + Trim(tmp(2)) + "-" + Trim(tmp(0)) + "-" + Trim(tmp(1))
    
    If flg Then
        Dim timenum As Long
        tmp_time = Split(tmp(3), ":")
        timenum = Val(tmp_time(0))
        If timenum < 9 Then
            hosei = hosei + 1
            chk = chk + " " + CStr(timenum + 12) + ":"
        Else
            chk = chk + " " + Trim(tmp_time(0)) + ":"
        End If
        chk = chk + Left(Trim(tmp_time(1)), 2) + ":00" + """"
    Else
        chk = chk + " " + Left(Trim(tmp(3)), 4) + ":00" + """"
    End If
    hosei = hosei + 1
    TransExe = TransExe + chk + Mid(inData, i + hosei, length - i + hosei)
    length = Len(TransExe)
End Function
Private Function oddchekc() As Boolean
    Dim temp As String
    Dim i As Long
    Dim count As Long: count = 0
    For i = 1 To length
        temp = Mid(data, i, 1)
        If temp = """" Then
            count = count + 1
        End If
    Next i
    If count Mod 2 = 0 Then
        oddchekc = False
    Else
        oddchekc = True
    End If
End Function
