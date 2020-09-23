VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Current parsing positions as indicated by |
'  |<foo>   =>  lPos = 0
'  <foo>|   =>  lPos = 5
'


' private variables =========================================================================================

Private pxDocument As XMLNode
Private plDepth As Long



' public properties =========================================================================================

Public Property Get XML() As String
  Dim strXML As String
  strXML = "<?xml" & pxDocument.Attributes.XML & "?>" & vbCrLf
  If Me.DocType <> "" Then strXML = strXML & "<!DOCTYPE " & Me.DocType & ">" & vbCrLf
  XML = strXML & pxDocument.Children.XML
End Property


Public Property Get Depth() As Long
  Depth = plDepth
End Property

Public Property Get Length() As Long
  Length = pxDocument.Length
End Property


Public Property Let DocType(ByRef rstrType As String)
'  Dim xNode As XMLNode
'  '
'  ' no checking here yet
'  '
'  If Not Me.Children.IsElement("DOCTYPE") Then
'    Set xNode = New XMLNode
'    With xNode
'      .NodeType = xntDOCTYPE
'      .Text = rstrType
'  pstrDocType = rstrType
End Property

Public Property Get DocType() As String
  If Me.Children.IsElement("DOCTYPE") Then
    DocType = Me.Children("DOCTYPE")
  End If
End Property


'
' document's node properties ----------------------------------------------------------------

Public Property Get Attributes() As XMLAttributes
  Set Attributes = pxDocument.Attributes
End Property

Public Property Get Children() As XMLNodes
  Set Children = pxDocument.Children
End Property

Public Property Get EnumChildrenByName(ByRef rstrName As String) As XMLNodes
  Set EnumChildrenByName = pxDocument.EnumChildrenByName(rstrName)
End Property

Public Property Get HasChildren() As Boolean
  HasChildren = pxDocument.HasChildren
End Property

Public Property Get First(Optional ByRef rstrName As String, Optional ByRef rstrID As String) As XMLNode
  Set First = pxDocument.First(rstrName, rstrID)
End Property



' public methods =========================================================================================

Public Sub LoadXML(ByRef rstrXML As String)
  Const pcstrErrSrc As String = "XMLDocument.LoadXML"
  Dim xNode As XMLNode
  Dim strTmp As String, lPos As Long, n As Long, lNextClose As Long, eType As XMLNodeTypes
  
  Set pxDocument = New XMLNode
  With pxDocument
    .NodeType = xntNode
    .Name = "xml-document-root"
  End With
  
  n = InStr(rstrXML, "<?xml ")
  If n < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Missing XML declaration"
  strTmp = Left$(rstrXML, n - 1)
  If XMLNormalizeWhitespace(strTmp) <> "" Then Err.Raise vbObjectError, pcstrErrSrc, "Non-whitespace before XML declaration"
  lPos = n + 6  ' <?xml ?|
  n = InStr(lPos, rstrXML, "?>")
  If n < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Invalid XML declaration"
  strTmp = Mid$(rstrXML, lPos, n - lPos)
  lPos = n + 1 ' ?>?|
  pxDocument.Attributes.Parse Trim$(strTmp)
  If Not pxDocument.Attributes.IsElement("version") Or Not pxDocument.Attributes.IsElement("encoding") Then Err.Raise vbObjectError, pcstrErrSrc, "Invalid XML declaration"
  If Not pxDocument.Attributes("version") = "1.0" Then Err.Raise vbObjectError, pcstrErrSrc, "Non-supported XML version"
  If Not pxDocument.Attributes("encoding") = "UTF-8" Then Err.Raise vbObjectError, pcstrErrSrc, "Non-supported encoding"
     
  '
  ' let's play some nasty trick
  '
  plDepth = pxDocument.eat(lPos, rstrXML & "</xml-document-root>")
  
  n = 0
  
  For Each xNode In pxDocument.Children
    eType = xNode.NodeType
    If eType = xntNode Or xNode.NodeType = xntEmpty Then n = n + 1
    If eType = xntCDATA Or eType = xntText Then Err.Raise vbObjectError, pcstrErrSrc, "Illegal node type outside root"
  Next
  
  If n > 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Multiple root nodes"
End Sub


Public Sub Load(ByRef rstrFile As String)
  Dim iFile As Integer, strXML As String, cData() As Byte, lLen As Long
  
  On Error GoTo catch
    iFile = FreeFile
    Open rstrFile For Binary Access Read As iFile
    lLen = LOF(iFile)
    If lLen > 0 Then
      ReDim cData(lLen - 1)
      Get #1, , cData()
    End If
    Close iFile
  On Error GoTo 0
  
  If lLen > 0 Then
    'strXML = String$(lLen, 0)
    'BlowUp2Unicode cData(0), lLen, StrPtr(strXML), Len(strXML)   ' api version
    BlowUp2Unicode cData(), strXML
  End If
  
  Erase cData()
  
  LoadXML strXML
  
  Exit Sub
  
catch:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub Save(ByRef rstrFile As String)
  Dim iFile As Integer, strXML As String, cData() As Byte, lLen As Long
  
  strXML = Me.XML
  lLen = Len(strXML)
  
  If lLen < 1 Then Exit Sub
  
  'ReDim cData(lLen - 1)
  'Shrink2Bytes StrPtr(strXML), lLen, cData(0), lLen  ' api version
  Shrink2Bytes strXML, cData()
  strXML = "" ' save mem :)
  
  On Error Resume Next
    Kill rstrFile
  On Error GoTo catch
    iFile = FreeFile
    Open rstrFile For Binary Access Write As iFile
    Put #1, , cData()
    Close iFile
  On Error GoTo 0
  
  Exit Sub
  
catch:
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub


Public Function Validate() As Boolean
  LoadXML Me.XML
  Validate = True
End Function


' event procedures =========================================================================================

Private Sub Class_Initialize()
  Set pxDocument = New XMLNode
  With pxDocument.Attributes
    .Add "version", "1.0"
    .Add "encoding", "UTF-8"
  End With
End Sub

