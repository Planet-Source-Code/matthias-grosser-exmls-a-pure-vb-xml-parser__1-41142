VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' public constants =========================================================================================

Public Enum XMLNodeTypes
  xntUnknown    ' undefined, default and illegal
  xntNode       ' real node, à la <node>...</node>
  xntEmpty      ' no contents node: <node/>
  xntText       ' pure text, like foo and bar in <p>foo<b>blablabla</b>bar</p>
  xntCDATA      ' CDATA(non-entity-encoded) node: <![CDATA[...]]>
  xntComment    ' <!--...-->
  xntDOCTYPE    ' <!DOCTYPE ..>
End Enum

' private constants =========================================================================================

Private Const clXMLIndentPerLevel As Long = 2


' private variables =========================================================================================

Private pxChildren As XMLNodes
Private pxAttributes As XMLAttributes

Private pstrName As String
Private pstrText As String
Private peNodeType As XMLNodeTypes


' public properties =========================================================================================

Public Property Let Name(ByRef rstrName As String)
  If Not XMLIsLegalElementName(rstrName) Or pstrName <> "" Then Err.Raise vbObjectError, "XMLNode.Name", "Name illegal or already set"
  pstrName = rstrName
End Property

Public Property Get Name() As String    ' must only be set for xntNode and xntEmpty nodes!
  If (peNodeType = xntEmpty Or peNodeType = xntNode) And pstrName = "" Then Err.Raise vbObjectError, "XMLNode.Name", "Required property not initialized"
  Name = pstrName
End Property


Public Property Get HasChildren() As Boolean
  HasChildren = (pxChildren.Count > 0)
End Property


Public Property Get NodeType() As XMLNodeTypes
  If peNodeType = xntUnknown Then Err.Raise vbObjectError, "XMLNode.NodeType", "Required property not initialized"
  NodeType = peNodeType
End Property

Public Property Let NodeType(ByVal eType As XMLNodeTypes)
  If peNodeType <> xntUnknown Then Err.Raise vbObjectError, "XMLNode.NodeType", "Node type illegal or already set"
  peNodeType = eType
End Property


Public Property Let Text(ByRef rstrText As String)
  If peNodeType = xntEmpty Then
    Err.Raise vbObjectError, "XMLNode.Text", "Empty node may not contain text"
  ElseIf peNodeType = xntCDATA Then
    pstrText = StringReplace(rstrText, "]]>", "] >")
  Else
    pstrText = rstrText
  End If
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
  Text = pstrText
End Property


Public Property Get Value() As String
  On Error Resume Next
  Value = pxAttributes("value").Value
End Property

Public Property Let Value(ByRef rstrValue As String)
  pxAttributes.Add "value", rstrValue
End Property

Public Property Get LngValue() As Long
  On Error GoTo catch
  LngValue = CLng(Value)
catch:
End Property

Public Property Get BoolValue() As Boolean
  On Error GoTo catch
  BoolValue = CBool(Value)
catch:
End Property


Public Property Get ID() As String
  On Error Resume Next
  ID = pxAttributes("ID")
End Property

Public Property Let ID(ByRef rstrID As String)
  pxAttributes.Add "ID", rstrID
End Property


Public Property Get Children() As XMLNodes
  Set Children = pxChildren
End Property


Public Property Get Attributes() As XMLAttributes
  'If peNodeType <> xntNode And peNodeType <> xntEmpty Then Err.Raise vbObjectError, "XMLNode.Attributes", "Attributes only legal for Node and Empty node types"
  Set Attributes = pxAttributes
End Property


Public Property Get EnumChildrenByName(ByRef rstrName As String) As XMLNodes
  Dim xNodes As XMLNodes, xNode As XMLNode
  Set xNodes = New XMLNodes
  For Each xNode In pxChildren
    If xNode.Name = rstrName Then xNodes.Add xNode
  Next
  Set EnumChildrenByName = xNodes
End Property


Public Property Get First(Optional ByRef rstrName As String, Optional ByRef rstrID As String) As XMLNode
  Dim xNode As XMLNode

  If rstrName = "" And rstrID = "" Then
    ' get first real node
    For Each xNode In pxChildren
      If xNode.NodeType = xntNode Or xNode.NodeType = xntEmpty Then Set First = xNode: Exit Property
    Next
    
  ElseIf rstrName <> "" And rstrID = "" Then
    ' get first named node
    For Each xNode In pxChildren
      If (xNode.NodeType = xntNode Or xNode.NodeType = xntEmpty) And xNode.Name = rstrName Then Set First = xNode: Exit Property
    Next
    
  ElseIf rstrName <> "" And rstrID <> "" Then
    ' get node by ID
    Set First = pxChildren(rstrName, rstrID)
    
  Else
    ' ID without name - illegal
    Err.Raise vbObjectError, "XMLDocument.First", "Node name required"
    
  End If
End Property


Public Property Get Length() As Long
  Length = 1 + pxChildren.Length
End Property


Friend Property Get niceXML(ByVal lLevel As Long) As String
  Dim xNode As XMLNode
  Dim strXML As String, strIndent As String
  
  strIndent = String$(lLevel * clXMLIndentPerLevel, " ")
  
  Select Case peNodeType
    Case xntNode
      strXML = "<" & Me.Name
      
    Case xntEmpty
      strXML = "<" & Me.Name
      
    Case xntText
      niceXML = UTF8Encode(XMLEntityEncode(pstrText))
      Exit Property
      
    Case xntCDATA
      niceXML = strIndent & "<![CDATA[" & UTF8Encode(pstrText) & "]]>" & vbCrLf
      Exit Property

    Case xntComment
      niceXML = strIndent & "<!--" & UTF8Encode(pstrText) & "-->" & vbCrLf
      Exit Property
      
    Case xntDOCTYPE
      niceXML = strIndent & "<!DOCTYPE " & pstrText & ">" & vbCrLf
      Exit Property

    Case Else
      Err.Raise vbObjectError, "XMLNode.XML", "Illegal Node Type"
      
  End Select
    
  '
  ' only xntNode and xntEmpty nodes here
  '
  strXML = strXML & pxAttributes.XML
  
  If peNodeType = xntEmpty Then
    niceXML = strIndent & strXML & "/>" & vbCrLf
    Exit Property
  End If
  
  ' only xntNode nodes here
  
  
  strXML = strXML & ">"   ' close open tag
  If Not Me.HasChildren Then
    niceXML = strIndent & strXML & UTF8Encode(XMLEntityEncode(pstrText)) & "</" & Me.Name & ">" & vbCrLf  ' close tag
    Exit Property
  End If
  
  strXML = strXML & vbCrLf
  'For Each xNode In Me.Children
  strXML = strXML & pxChildren.niceXML(lLevel + 1)   ' add contents
  'Next
  niceXML = strIndent & strXML & strIndent & "</" & Me.Name & ">" & vbCrLf  ' close tag
End Property


Public Property Get XML() As String
  XML = niceXML(0)
End Property


' private methods =========================================================================================

'
' main parsing takes place here
' requires: Name property defined
'
' this is only called for OPENING nodes!
'
' inspired by the pacman game :-)
'
Friend Function eat(ByRef rlPos As Long, ByRef rstrXML As String) As Long
  Const pcstrErrSrc As String = "XMLNode.eat"
  Dim xNode As XMLNode
  Dim n As Long, m As Long, strTmp As String, strName As String, lLevel As Long
  
  pxChildren.Clear
  
  If pstrName = "" Then Err.Raise vbObjectError, pcstrErrSrc, "Node name required"
  
  Do
  
    ' rlPos: >|
    n = InStr(rlPos + 1, rstrXML, "<")
    If n = 0 Then Err.Raise vbObjectError, pcstrErrSrc, "Node not closed"
    
    '
    ' text zwischen anfang und hier jetzt schon merken
    '
    strTmp = Mid$(rstrXML, rlPos + 1, n - rlPos - 1)
    If XMLAnyTags(strTmp) Then Err.Raise vbObjectError, pcstrErrSrc, "Illegal encoding"
    
    
    If Mid$(rstrXML, n + 1, 1) = "/" Then ' close, MUST be me
      m = InStr(n + 2, rstrXML, ">")
      If m < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Tag not closed"
      If XMLRemoveWhitespace(Mid$(rstrXML, n + 2, m - n - 2)) <> pstrName Then Err.Raise vbObjectError, pcstrErrSrc, "Badly nested tags"
      
      If Me.HasChildren Then    ' add text node
        If XMLRemoveWhitespace(strTmp) <> "" Then
          If Left$(strTmp, 2) = vbCrLf Then strTmp = Mid$(strTmp, 3)  ' preserve whitespace, avoid file growth
          Set xNode = New XMLNode
          With xNode
            .NodeType = xntText
            .Text = UTF8Decode(XMLEntityDecode(strTmp))
          End With
          pxChildren.Add xNode
        End If
      Else
        pstrText = UTF8Decode(XMLEntityDecode(strTmp))
      End If
      
      
      'fertig
      ' halt vorher ggef textnode von oben noch mitnehmen
      ' nur wenn node hasChildren!
      
      
      rlPos = m
      ' RAUSSPRINGEN?!
      Exit Do
      
    End If
        
    '
    ' there are more children to come -> create a text node?
    '
    If XMLRemoveWhitespace(strTmp) <> "" Then
      If Left$(strTmp, 2) = vbCrLf Then strTmp = Mid$(strTmp, 3)  ' preserve whitespace, avoid file growth
      Set xNode = New XMLNode
      With xNode
        .NodeType = xntText
        .Text = UTF8Decode(XMLEntityDecode(strTmp))
      End With
      pxChildren.Add xNode
    End If
    
    '
    ' continue finding out node type
    '
    If Mid$(rstrXML, n + 1, 1) = "!" Then  ' comment or cdata
      If Mid$(rstrXML, n + 2, 2) = "--" Then ' comment
        m = InStr(n + 4, rstrXML, "-->")
        If m < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Comment not closed"
        Set xNode = New XMLNode
        With xNode
          .NodeType = xntComment
          .Text = UTF8Decode(Mid$(rstrXML, n + 4, m - n - 4))
        End With
        pxChildren.Add xNode
        rlPos = m + 3
        
      ElseIf Mid$(rstrXML, n + 2, 7) = "[CDATA[" Then
        m = InStr(n + 9, rstrXML, "]]>")
        If m < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "CDATA node not closed"
        Set xNode = New XMLNode
        With xNode
          .NodeType = xntCDATA
          .Text = UTF8Decode(Mid$(rstrXML, n + 9, m - n - 9))   ' no entityencoding for CDATA
        End With
        pxChildren.Add xNode
        rlPos = m + 2
        
      ElseIf Mid$(rstrXML, n + 2, 8) = "DOCTYPE " Then
        m = InStr(n + 10, rstrXML, ">")
        If m < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "DOCTYPE node not closed"
        '
        ' check if valid DOCTYPE here!
        '
        Set xNode = New XMLNode
        With xNode
          .NodeType = xntDOCTYPE
          .Text = Mid$(rstrXML, n + 10, m - n - 10)
        End With
        pxChildren.Add xNode
        rlPos = m + 1
        
      Else  ' unsupported
        Err.Raise vbObjectError, pcstrErrSrc, "Unsupported <! node type"
      
      End If
      
    Else  ' new opening or empty node
      m = InStr(n + 1, rstrXML, ">")
      If m < 1 Then Err.Raise vbObjectError, pcstrErrSrc, "Tag not closed @ " & n
      
      strTmp = Mid$(rstrXML, n + 1, m - n - 1)   ' <contents>
      
      If Right$(strTmp, 1) = "/" Then   ' xntEmpty node
        strTmp = Left$(strTmp, Len(strTmp) - 1)
        strName = FirstArg(strTmp, " ", , True)
        If Not XMLIsLegalElementName(strName) Then Err.Raise vbObjectError, pcstrErrSrc, "Illegal element name"
        strTmp = Mid$(strTmp, Len(strName) + 1)
        Set xNode = New XMLNode
        With xNode
          .NodeType = xntEmpty
          .Name = strName
          .Attributes.Parse Trim$(strTmp)
        End With
        pxChildren.Add xNode
        rlPos = m
        
      Else  ' OPENING node
        strName = FirstArg(strTmp, " ", , True)
        If Not XMLIsLegalElementName(strName) Then Err.Raise vbObjectError, pcstrErrSrc, "Illegal element name"
        strTmp = Mid$(strTmp, Len(strName) + 1)
        Set xNode = New XMLNode
        With xNode
          .NodeType = xntNode
          .Name = strName
          .Attributes.Parse Trim$(strTmp)
        End With
        pxChildren.Add xNode
        rlPos = m ' >|
      
        m = xNode.eat(rlPos, rstrXML)
        If m > lLevel Then lLevel = m
      
      End If
    End If
  
  Loop While rlPos < Len(rstrXML)
  
  eat = lLevel + 1
    
End Function


' event procedures =========================================================================================

Private Sub Class_Initialize()
  Set pxChildren = New XMLNodes
  Set pxAttributes = New XMLAttributes
End Sub
