Attribute VB_Name = "eXMLHeaders"
Option Explicit

' eXMLs 1.0
'
' Created & (c) 2002 Matthias Grosser
'
' You may use this code for non-commercial purposes as you like,
' but should never ever suppose it being completely free of bugs.
' On discovering any, please mail them to pluto@brain-killer.org.
'
' http://pluto.brain-killer.org/dev/exmls/
'


' api declarations =========================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


' global utility functions =========================================================================================

'
'&lt; produces the left angle bracket, <
'&gt; produces the right angle bracket, >
'&amp; produces the ampersand, &
'&apos; produces a single quote character (an apostrophe), '
'&quot; produces a double quote character, "
'

Public Function XMLEntityEncode(ByVal strData As String) As String
  strData = StringReplace(strData, "&", "&amp;")
  strData = StringReplace(strData, "<", "&lt;")
  strData = StringReplace(strData, ">", "&gt;")
  strData = StringReplace(strData, """", "&quot;")
  strData = StringReplace(strData, "'", "&apos;")
  XMLEntityEncode = strData
End Function


Public Function XMLEntityDecode(ByVal strData As String) As String
  strData = StringReplace(strData, "&apos;", "'")
  strData = StringReplace(strData, "&quot;", """")
  strData = StringReplace(strData, "&gt;", ">")
  strData = StringReplace(strData, "&lt;", "<")
  strData = StringReplace(strData, "&amp;", "&")
  XMLEntityDecode = strData
End Function


Public Function XMLNormalizeWhitespace(ByVal strXML As String) As String
  Dim lPrevLen As Long
  
  strXML = StringReplace(strXML, vbLf, " ")
  strXML = Trim$(StringReplace(strXML, vbCr, " "))
  Do
    lPrevLen = Len(strXML)
    strXML = StringReplace(strXML, "  ", " ")
  Loop While lPrevLen > Len(strXML)
  XMLNormalizeWhitespace = strXML
End Function


Public Function XMLRemoveWhitespace(ByVal strXML As String) As String
  strXML = StringReplace(strXML, vbLf, "")
  strXML = Trim$(StringReplace(strXML, vbCr, ""))
  strXML = StringReplace(strXML, " ", "")
  XMLRemoveWhitespace = strXML
End Function


Public Function XMLIsLegalElementName(ByRef rstrName As String) As Boolean
  If rstrName = "" Then Exit Function
  If InStr(rstrName, " ") > 0 Then Exit Function
  If InStr(rstrName, vbLf) > 0 Then Exit Function
  If InStr(rstrName, vbCr) > 0 Then Exit Function
  If InStr(rstrName, "/") > 0 Then Exit Function
  If InStr(rstrName, """") > 0 Then Exit Function
  If InStr(rstrName, "&") > 0 Then Exit Function
  If InStr(rstrName, "<") > 0 Then Exit Function
  If InStr(rstrName, ">") > 0 Then Exit Function
  If InStr(rstrName, "'") > 0 Then Exit Function
  XMLIsLegalElementName = True
End Function


Public Function XMLAnyTags(ByRef rstrData As String) As Boolean
  XMLAnyTags = True
  If InStr(rstrData, "<") > 0 Then Exit Function
  If InStr(rstrData, ">") > 0 Then Exit Function
  XMLAnyTags = False
End Function


Public Function XMLIsLegalEncoding(ByRef rstrData As String, Optional ByRef rstrQuoteChar As String = """") As Boolean
  If InStr(rstrData, rstrQuoteChar) > 0 Then Exit Function
  If InStr(rstrData, "<") > 0 Then Exit Function
  If InStr(rstrData, ">") > 0 Then Exit Function
  XMLIsLegalEncoding = True
End Function


Public Function CXMLBool(ByVal bValue As Boolean) As String
  If bValue Then CXMLBool = "true" Else CXMLBool = "false"
End Function


Public Function UTF8Encode(ByRef rstrUnicode As String) As String
  Dim iData() As Integer, lLen As Long, n As Long, iUtf() As Integer, i As Long, strUtf As String
  lLen = Len(rstrUnicode)
  If lLen = 0 Then Exit Function
  ReDim iData(lLen - 1)
  ReDim iUtf(lLen * 3 - 1)
  CopyMemory iData(0), ByVal StrPtr(rstrUnicode), lLen * 2
  For n = 0 To lLen - 1
    If (iData(n) And &HFF80) = 0 Then 'iData(n) < &H80 And Not iData(n) < 0 Then
      iUtf(i) = iData(n)
      i = i + 1
    ElseIf (iData(n) And &HF800) = 0 Then 'iData(n) < &H800 And Not iData(n) < 0 Then
      iUtf(i) = &HFF And (&HC0 Or iRSH(iData(n), 6))
      i = i + 1
      iUtf(i) = &HFF And (&H80 Or iData(n) And &H3F)
      i = i + 1
    Else
      iUtf(i) = &HFF And (&HE0 Or iRSH(iData(n), 12))
      i = i + 1
      iUtf(i) = &HFF And (&H80 Or iRSH(iData(n), 6) And &H3F)
      i = i + 1
      iUtf(i) = &HFF And (&H80 Or iData(n) And &H3F)
      i = i + 1
    End If
  Next
  strUtf = String$(i, 0)
  CopyMemory ByVal StrPtr(strUtf), iUtf(0), i * 2
  UTF8Encode = strUtf
End Function


Public Function UTF8Decode(ByRef rstrUtf As String) As String
  Dim iUtf() As Integer, iUnicode() As Integer, n As Long, i As Long, strUnicode As String, lLen As Long
  lLen = Len(rstrUtf)
  If lLen = 0 Then Exit Function
  ReDim iUtf(lLen - 1)
  ReDim iUnicode(lLen - 1)
  CopyMemory iUtf(0), ByVal StrPtr(rstrUtf), lLen * 2
  
  On Error GoTo catch   ' catch index out of range errors caused by illegal utf sequences
  
  For n = 0 To lLen - 1
    If iUtf(n) > &HEF Then
      GoTo catch
    ElseIf (iUtf(n) And &HF0) = &HE0 Then   ' 3 byte seq.
      If (iUtf(n + 1) And &HC0) <> &H80 Or (iUtf(n + 2) And &HC0) <> &H80 Then GoTo catch
      iUnicode(i) = (iLSH(iUtf(n) And &HF, 12) Or iLSH(iUtf(n + 1) And &H3F, 6) Or (iUtf(n + 2) And &H3F))
      i = i + 1
      n = n + 2
    ElseIf (iUtf(n) And &HE0) = &HC0 Then   ' 2 byte seq.
      If (iUtf(n + 1) And &HC0) <> &H80 Then GoTo catch
      iUnicode(i) = (iLSH(iUtf(n) And &H1F, 6) Or (iUtf(n + 1) And &H3F))
      i = i + 1
      n = n + 1
    Else   ' 1 byte char
      iUnicode(i) = iUtf(n)
      i = i + 1
    End If
  Next
  strUnicode = String$(i, 0)
  CopyMemory ByVal StrPtr(strUnicode), iUnicode(0), i * 2
  UTF8Decode = strUnicode
  
  Exit Function

catch:
  UTF8Decode = rstrUtf
End Function

Private Function iRSH(ByVal i As Integer, s As Integer) As Integer
  Dim k As Long
  On Error GoTo catch
  If s = 0 Then
    iRSH = i
  Else
    k = i And &HFFFF&
    iRSH = k \ (2& ^ s)
  End If
catch:
End Function

Private Function iLSH(ByVal i As Integer, s As Integer) As Integer
  Dim k As Long, iRetVal As Integer
  On Error GoTo catch
  If s = 0 Then
    iLSH = i
  Else
    k = i And &HFFFF&
    k = (k * (2& ^ s)) And &HFFFF&
    If k >= &H8000& Then
      iLSH = k - 65536
    Else
      iLSH = k
    End If
  End If
catch:
End Function


Public Sub BlowUp2Unicode(ByRef rcData() As Byte, ByRef rstrOut As String)
  Dim n As Long, m As Long, iData() As Integer
  On Error GoTo catch
    m = UBound(rcData())
  On Error GoTo 0
  ReDim iData(m)
  For n = 0 To m
    iData(n) = rcData(n)
  Next
  rstrOut = String$(m + 1, 0)
  CopyMemory ByVal StrPtr(rstrOut), iData(0), Len(rstrOut) * 2
  Exit Sub
catch:
  rstrOut = ""
End Sub


Public Sub Shrink2Bytes(ByRef rstrIn As String, ByRef rcData() As Byte)
  Dim iData() As Integer, m As Long, n As Long
  m = Len(rstrIn)
  If m = 0 Then Erase rcData(): Exit Sub
  ReDim iData(m - 1)
  CopyMemory iData(0), ByVal StrPtr(rstrIn), m * 2
  On Error Resume Next
    For n = 0 To m
      rcData(n) = iData(n)
    Next
  On Error GoTo 0
End Sub


Public Function StringReplace(ByVal strIn As String, ByRef rstrReplaceThis As String, ByRef rstrByThis As String, Optional ByVal eMethod As VbCompareMethod = vbBinaryCompare) As String
  Dim lPos As Long, lLength As Long, lByLength As Long
  lLength = Len(rstrReplaceThis)
  lByLength = Len(rstrByThis)
  If lLength = 0 Then StringReplace = strIn: Exit Function
  lPos = InStr(1, strIn, rstrReplaceThis, eMethod)
  While lPos > 0
    strIn = Left$(strIn, lPos - 1) & rstrByThis & Mid$(strIn, lPos + lLength)
    lPos = InStr(lPos + lByLength, strIn, rstrReplaceThis, eMethod)
  Wend
  StringReplace = strIn
End Function


Public Function FirstArg(ByRef rstrIn As String, Optional ByRef rstrDelimiter As String = "#", Optional ByVal lOffset As Long = 0, Optional ByVal bAllIfNone As Boolean = True) As String
  Dim lHash As Long

  lHash = InStr(lOffset + 1, rstrIn, rstrDelimiter)
  If lHash > 0 Then
    FirstArg = Mid$(rstrIn, lOffset + 1, lHash - lOffset - 1)
  Else
    If bAllIfNone Then FirstArg = Mid$(rstrIn, lOffset + 1)
  End If
End Function

