Attribute VB_Name = "modJson"
Option Explicit
'===============================================================
' modJson - JSON lesen und schreiben (ohne externe Verweise)
' Objekte  -> Scripting.Dictionary (Reihenfolge bleibt erhalten)
' Arrays   -> Collection (1-basiert)
' Zahlen   -> Long oder Double, true/false -> Boolean, null -> Null
'===============================================================
Private m_s As String
Private m_p As Long
Private m_len As Long

Public Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0   ' Binaer: Gross-/Kleinschreibung wird unterschieden
    Set NewDict = d
End Function

Public Function NewList() As Collection
    Set NewList = New Collection
End Function

'---------------------------------------------------------------
' Lesen
'---------------------------------------------------------------
Public Function JsonParse(ByVal s As String) As Variant
    Dim r As Variant
    m_s = s
    m_len = Len(s)
    m_p = 1
    If m_len > 0 Then
        If Mid$(m_s, 1, 1) = ChrW(&HFEFF) Then m_p = 2   ' BOM ueberspringen
    End If
    ParseValue r
    SkipWs
    If m_p <= m_len Then JsonError "Unerwartete Zeichen nach dem Ende des JSON-Dokuments"
    If IsObject(r) Then
        Set JsonParse = r
    Else
        JsonParse = r
    End If
End Function

Private Sub JsonError(ByVal msg As String)
    Err.Raise vbObjectError + 1001, "modJson", "JSON-Fehler: " & msg & " (Position " & m_p & ")"
End Sub

Private Sub SkipWs()
    Dim c As String
    Do While m_p <= m_len
        c = Mid$(m_s, m_p, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            m_p = m_p + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ParseValue(ByRef result As Variant)
    Dim c As String
    SkipWs
    If m_p > m_len Then JsonError "Unerwartetes Ende des Dokuments"
    c = Mid$(m_s, m_p, 1)
    Select Case c
        Case "{"
            Set result = ParseObject()
        Case "["
            Set result = ParseArray()
        Case """"
            result = ParseString()
        Case "t"
            ExpectWord "true"
            result = True
        Case "f"
            ExpectWord "false"
            result = False
        Case "n"
            ExpectWord "null"
            result = Null
        Case Else
            result = ParseNumber()
    End Select
End Sub

Private Sub ExpectWord(ByVal w As String)
    If Mid$(m_s, m_p, Len(w)) <> w Then JsonError "Ungueltiges Literal"
    m_p = m_p + Len(w)
End Sub

Private Function ParseObject() As Object
    Dim d As Object
    Set d = NewDict()
    m_p = m_p + 1   ' {
    SkipWs
    If Mid$(m_s, m_p, 1) = "}" Then
        m_p = m_p + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        ParseMember d
        SkipWs
        Select Case Mid$(m_s, m_p, 1)
            Case ","
                m_p = m_p + 1
            Case "}"
                m_p = m_p + 1
                Exit Do
            Case Else
                JsonError "',' oder '}' erwartet"
        End Select
    Loop
    Set ParseObject = d
End Function

Private Sub ParseMember(ByVal d As Object)
    Dim key As String, v As Variant
    SkipWs
    If Mid$(m_s, m_p, 1) <> """" Then JsonError "Schluessel (Text in Anfuehrungszeichen) erwartet"
    key = ParseString()
    SkipWs
    If Mid$(m_s, m_p, 1) <> ":" Then JsonError "':' erwartet"
    m_p = m_p + 1
    ParseValue v
    If IsObject(v) Then
        Set d(key) = v
    Else
        d(key) = v
    End If
End Sub

Private Function ParseArray() As Collection
    Dim col As Collection
    Set col = New Collection
    m_p = m_p + 1   ' [
    SkipWs
    If Mid$(m_s, m_p, 1) = "]" Then
        m_p = m_p + 1
        Set ParseArray = col
        Exit Function
    End If
    Do
        ParseElement col
        SkipWs
        Select Case Mid$(m_s, m_p, 1)
            Case ","
                m_p = m_p + 1
            Case "]"
                m_p = m_p + 1
                Exit Do
            Case Else
                JsonError "',' oder ']' erwartet"
        End Select
    Loop
    Set ParseArray = col
End Function

Private Sub ParseElement(ByVal col As Collection)
    Dim v As Variant
    ParseValue v
    col.Add v
End Sub

Private Function ParseString() As String
    Dim sb As String, c As String, q As Long, bs As Long
    m_p = m_p + 1   ' oeffnendes Anfuehrungszeichen
    Do
        q = InStr(m_p, m_s, """")
        If q = 0 Then JsonError "Text nicht abgeschlossen"
        bs = InStr(m_p, m_s, "\")
        If bs = 0 Or q < bs Then
            sb = sb & Mid$(m_s, m_p, q - m_p)
            m_p = q + 1
            Exit Do
        End If
        sb = sb & Mid$(m_s, m_p, bs - m_p)
        m_p = bs + 1
        c = Mid$(m_s, m_p, 1)
        Select Case c
            Case """": sb = sb & """"
            Case "\": sb = sb & "\"
            Case "/": sb = sb & "/"
            Case "b": sb = sb & vbBack
            Case "f": sb = sb & Chr$(12)
            Case "n": sb = sb & vbLf
            Case "r": sb = sb & vbCr
            Case "t": sb = sb & vbTab
            Case "u"
                sb = sb & ChrW(CLng("&H" & Mid$(m_s, m_p + 1, 4)))
                m_p = m_p + 4
            Case Else
                JsonError "Ungueltige Escape-Sequenz"
        End Select
        m_p = m_p + 1
    Loop
    ParseString = sb
End Function

Private Function ParseNumber() As Variant
    Dim st As Long, c As String, tok As String, isInt As Boolean
    st = m_p
    isInt = True
    Do While m_p <= m_len
        c = Mid$(m_s, m_p, 1)
        If InStr("+-0123456789.eE", c) > 0 Then
            If c = "." Or c = "e" Or c = "E" Then isInt = False
            m_p = m_p + 1
        Else
            Exit Do
        End If
    Loop
    If m_p = st Then JsonError "Unerwartetes Zeichen '" & Mid$(m_s, st, 1) & "'"
    tok = Mid$(m_s, st, m_p - st)
    If isInt And Len(tok) <= 9 Then
        ParseNumber = CLng(Val(tok))
    Else
        ParseNumber = Val(tok)
    End If
End Function

'---------------------------------------------------------------
' Schreiben
'---------------------------------------------------------------
Public Function JsonStringify(ByVal v As Variant, Optional ByVal pretty As Boolean = True) As String
    Dim sb As String
    WriteValue sb, v, 0, pretty
    JsonStringify = sb
End Function

Private Sub WriteValue(ByRef sb As String, ByVal v As Variant, ByVal level As Long, ByVal pretty As Boolean)
    Dim k As Variant, i As Long, first As Boolean, n As Long, lb As Long, ub As Long
    If IsObject(v) Then
        If v Is Nothing Then
            sb = sb & "null"
            Exit Sub
        End If
        Select Case TypeName(v)
            Case "Dictionary"
                If v.Count = 0 Then
                    sb = sb & "{}"
                    Exit Sub
                End If
                sb = sb & "{"
                first = True
                For Each k In v.Keys
                    If Not first Then sb = sb & ","
                    first = False
                    If pretty Then sb = sb & vbLf & Space$((level + 1) * 2)
                    sb = sb & Quote(CStr(k)) & ":"
                    If pretty Then sb = sb & " "
                    WriteValue sb, v(k), level + 1, pretty
                Next k
                If pretty Then sb = sb & vbLf & Space$(level * 2)
                sb = sb & "}"
            Case "Collection"
                n = v.Count
                If n = 0 Then
                    sb = sb & "[]"
                    Exit Sub
                End If
                sb = sb & "["
                For i = 1 To n
                    If i > 1 Then sb = sb & ","
                    If pretty Then sb = sb & vbLf & Space$((level + 1) * 2)
                    WriteValue sb, v(i), level + 1, pretty
                Next i
                If pretty Then sb = sb & vbLf & Space$(level * 2)
                sb = sb & "]"
            Case Else
                JsonError "Objekttyp wird nicht unterstuetzt: " & TypeName(v)
        End Select
    ElseIf VarType(v) >= vbArray Then
        lb = LBound(v)
        ub = UBound(v)
        If ub < lb Then
            sb = sb & "[]"
            Exit Sub
        End If
        sb = sb & "["
        For i = lb To ub
            If i > lb Then sb = sb & ","
            If pretty Then sb = sb & vbLf & Space$((level + 1) * 2)
            WriteValue sb, v(i), level + 1, pretty
        Next i
        If pretty Then sb = sb & vbLf & Space$(level * 2)
        sb = sb & "]"
    Else
        Select Case VarType(v)
            Case vbNull, vbEmpty
                sb = sb & "null"
            Case vbBoolean
                If v Then
                    sb = sb & "true"
                Else
                    sb = sb & "false"
                End If
            Case vbString
                sb = sb & Quote(CStr(v))
            Case vbDate
                sb = sb & Quote(Format$(v, "yyyy-mm-dd\Thh:nn:ss"))
            Case vbInteger, vbLong, vbByte, vbSingle, vbDouble, vbCurrency, vbDecimal
                sb = sb & Trim$(Str$(v))
            Case Else
                sb = sb & Quote(CStr(v))
        End Select
    End If
End Sub

Private Function Quote(ByVal s As String) As String
    Dim i As Long, c As Long, ch As String, out As String
    out = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        Select Case c
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case 0 To 31: out = out & "\u" & Right$("0000" & Hex$(c), 4)
            Case Else: out = out & ch
        End Select
    Next i
    Quote = out & """"
End Function

'---------------------------------------------------------------
' Kleine Zugriffshilfen
'---------------------------------------------------------------
Public Function JStr(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    If VarType(d(key)) = vbBoolean Then
        If d(key) Then JStr = "true" Else JStr = "false"
        Exit Function
    End If
    JStr = CStr(d(key))
End Function

Public Function JBool(ByVal d As Object, ByVal key As String) As Boolean
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If VarType(d(key)) = vbBoolean Then JBool = d(key)
End Function
