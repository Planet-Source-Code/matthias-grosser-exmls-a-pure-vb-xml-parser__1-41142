VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' private variables =========================================================================================

Private pcolAttributes As Collection


' public properties =========================================================================================

Public Property Get XML() As String
  Dim xAttribute As XMLAttribute, strXML As String
  
  For Each xAttribute In pcolAttributes
    strXML = strXML & " " & xAttribute.XML
  Next
  
  XML = strXML
End Property


' public methods =========================================================================================

'
' Attributes _always_ have an ID, so we use string-only keys.
'
Public Function Add(ByRef rstrID As String, Optional ByRef rstrValue As String = "") As XMLAttribute
  Dim xAttribute As XMLAttribute
  If rstrID = "" Then Err.Raise vbObjectError, "XMLAttribute.Add", "Attribute ID required"
  If Me.IsElement(rstrID) Then
    Set xAttribute = pcolAttributes(rstrID)   ' attribute already here, overwrite value
    xAttribute.Value = rstrValue
  Else
    Set xAttribute = New XMLAttribute         ' create new
    xAttribute.ID = rstrID
    xAttribute.Value = rstrValue
    pcolAttributes.Add xAttribute, rstrID
  End If
  Set Add = xAttribute
End Function



' receives [] in <foo[----]>
'
Public Function Parse(ByRef rstrXML As String) As Boolean
  Const pcstrErrSrc As String = "XMLAttributes.Parse"
  Dim xAttribute As XMLAttribute
  Dim strID As String, strData As String, strValue As String, lPos As Long, lFirstQuote As Long, lFirstApos As Long, strQuoteChar As String
  
  Me.Clear
  
  While lPos < Len(rstrXML)
    strID = FirstArg(rstrXML, "=", lPos, False)
    lPos = lPos + Len(strID) + 2   ' offset: ="|
    strID = LTrim$(strID)
    If Not XMLIsLegalElementName(strID) Then Err.Raise vbObjectError, pcstrErrSrc, "Illegal element name: " & strID
    lFirstQuote = InStr(lPos, rstrXML, """")
    lFirstApos = InStr(lPos, rstrXML, "'")
    If lFirstApos < 1 And lFirstQuote > 0 Or lFirstQuote < lFirstApos Then
      strQuoteChar = """"
    ElseIf lFirstApos > 0 And lFirstQuote < 1 Or lFirstQuote > lFirstApos Then
      strQuoteChar = "'"
    Else
      Err.Raise vbObjectError, pcstrErrSrc, "Malformed quotation"
    End If
    
    
    'If Mid$(rstrXML, lPos, 1) <> """" Then Err.Raise vbObjectError, pcstrErrSrc, "Malformed quotation"
    strValue = FirstArg(rstrXML, strQuoteChar, lPos, True)
    lPos = lPos + Len(strValue) + 1
    If Not XMLIsLegalEncoding(strValue, strQuoteChar) Then Err.Raise vbObjectError, pcstrErrSrc, "Bad attribute value encoding"
    
    strValue = UTF8Decode(XMLEntityDecode(strValue))
    
    If Me.IsElement(strID) Then
      Err.Raise vbObjectError, pcstrErrSrc, "Duplicate attributes"
      'Set xAttribute = pcolAttributes(strID)
      'xAttribute.Value = strValue
    Else
      Me.Add strID, strValue
    End If
  Wend
  
  Parse = True
End Function


Public Function Count() As Long
  Count = pcolAttributes.Count
End Function


Public Sub Remove(ByRef rstrKey As String)
  pcolAttributes.Remove rstrKey
End Sub


Public Sub Clear()
  Dim n As Long, m As Long
  n = pcolAttributes.Count
  For m = n To 1 Step -1
    pcolAttributes.Remove m       ' this is actually faster than Set pcolAttributes = New Collection
  Next
End Sub


Public Function IsElement(ByRef rstrKey As String) As Boolean
  Dim xAttribute As XMLAttribute
  On Error GoTo catch
  Set xAttribute = pcolAttributes(rstrKey)
  IsElement = True
  Exit Function
catch:
End Function


Public Function Item(ByRef rstrKey As String) As XMLAttribute
Attribute Item.VB_UserMemId = 0
  Set Item = pcolAttributes(rstrKey)
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = pcolAttributes.[_NewEnum]
End Function


' event procedures =========================================================================================

Private Sub Class_Initialize()
  Set pcolAttributes = New Collection
End Sub

