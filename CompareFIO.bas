Attribute VB_Name = "Module1"
Public Function CompareFIO(ByVal str_1 As String, Optional ByVal str_2 As String = "") As String
    ' Params
    Dim v As String
    Dim c As String
    Dim s1, s2 As String
    Dim TextStrng As String
    Dim Result_1() As String
    Dim Result_2() As String
    Dim WordCount_1 As Integer
    Dim WordCount_2 As Integer
    Dim canLink As Boolean
    
    canLink = True
    matchWordCount = 0
    
    
    ' Tekshirish
    
    
    TextStrng = "The Quick Brown Fox Jumps Over The Lazy Dog"
    
    Result_1() = Split(UCase(str_1))
    Result_2() = Split(UCase(str_2))
    
    WordCount_1 = UBound(Result_1()) + 1
    WordCount_2 = UBound(Result_2()) + 1
    
    Debug.Print 8
    
    ' MsgBox CStr(97)
    
    v = CStr(WordCount_1) & " " & CStr(WordCount_2)
    
    v = v & "-----"
    
    For i1 = 0 To WordCount_1 - 1
        If CInt(Len(Result_1(i1))) <= 1 Then
            canLink = False
            'MsgBox CStr("01--" & Len(Result_1(i1)))
        Else
            'MsgBox CStr("11--" & Len(Result_1(i1)))
        End If
        v = v & "-" & CStr(Len(Result_1(i1)))
    Next i1
    
    For i = WordCount_1 To 4
    Next i
    
    v = v & "---"
    
    For i2 = 0 To WordCount_2 - 1
        If Len(Result_2(i2)) <= 1 Then
            canLink = False
            'MsgBox CStr("02--" & Len(Result_2(i2)))
        Else
            'MsgBox CStr("12--" & Len(Result_2(i2)))
        End If
        v = v & "-" & CStr(Len(Result_2(i2)))
    Next i2

    
    If canLink Then
        matchWordCount = 0
        For i = 1 To WordCount_1
            s1 = Simplify(Clear2(Result_1(i - 1)))
            For j = 1 To WordCount_2
                s2 = Simplify(Clear2(Result_2(j - 1)))
                If s1 = s2 Then
                    matchWordCount = matchWordCount + 1
                End If
            Next j
        Next i
    End If
    
    v = v & "--- --" & CStr(matchWordCount)
    
    If matchWordCount < 2 Then
        canLink = False
    End If
        
    If canLink Then c = "+" Else c = "-"
    v = c & " " & v
    
    ' v = Result_1(0) + " -9-9---888 " + Result_1(1)
    
    'MsgBox "yutyu"
    
    'v = Clear2("qwe321'F")
    'v = "fgsdfg2"
    CompareFIO = v
End Function

Public Function Clear2(s As String) As String
    'Xarflardan boshqa belgilarni olib tashlaydi
    Dim mLE As String 'Lotin alfaviti uchun (massiv letter english)
    Dim mLR As String 'Kirill alfaviti uchun (massiv letter russian)
    Dim newS As String 'Hosil bo'lgan yangi satr uchun
    Dim isLetter As Boolean
    
    
    mLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    mLR = "ÀÁÂÃÄÅ¨ÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞß"
    
    newS = ""
    
    For i = 1 To Len(s)
        tChar = Mid(s, i, 1)
        isLetter = False
        For e = 1 To Len(mLE)
            If UCase(tChar) = UCase(Mid(mLE, e, 1)) Then
                isLetter = True
                Exit For
            End If
        Next e
        For r = 1 To Len(mLR)
            If UCase(tChar) = UCase(Mid(mLR, r, 1)) Then
                isLetter = True
                Exit For
            End If
        Next r
        If isLetter Then
            newS = newS & tChar
        End If
    Next i
    
    'MsgBox newS
    
    Clear2 = newS
    
End Function

Public Function Simplify(s As String) As String
    Dim tS As String
    tS = s
    
    tS = Replace(UCase(tS), UCase("ya"), UCase("a"))
    tS = Replace(UCase(tS), UCase("ye"), UCase("e"))
    tS = Replace(UCase(tS), UCase("kh"), UCase("x"))
    tS = Replace(UCase(tS), UCase("dj"), UCase("j"))
    
    tS = Replace(UCase(tS), UCase("h"), UCase("x"))
    tS = Replace(UCase(tS), UCase("q"), UCase("k"))
    
    Simplify = tS
End Function

