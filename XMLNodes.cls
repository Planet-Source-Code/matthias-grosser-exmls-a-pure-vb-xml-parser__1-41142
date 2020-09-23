VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private variables =========================================================================================

Private pcolNodes As Collection


' public properties =========================================================================================

Public Property Get XML() As String
  XML = niceXML(0)
End Property

Friend Property Get niceXML(ByVal lLevel As Long) As String
  Dim xNode As XMLNode, strXML As String
  For Each xNode In pcolNodes
    strXML = strXML & xNode.niceXML(lLevel) ' & vbCrLf & vbCrLf
  Next
  niceXML = strXML
End Property


Public Property Get Length() As Long
  Dim xNode As XMLNode, lLength As Long
  For Each xNode In pcolNodes
    lLength = lLength + xNode.Length
  Next
  Length = lLength
End Property


' public methods =========================================================================================

Public Sub Add(ByVal xNode As XMLNode) ', Optional ByRef rstrKey As String)
  Dim strID As String

  On Error Resume Next
    strID = xNode.Attributes("ID")
  On Error GoTo 0
  If strID <> "" Then
    If xNode.Name = "" Then Err.Raise vbObjectError, "XMLNodes.Add", "ID'ed nodes must have a name"
    On Error GoTo catch
      pcolNodes.Add xNode, xNode.Name & " " & strID   ' name first, ID may contain spaces
    On Error GoTo 0
  Else
    pcolNodes.Add xNode
  End If
  
  Exit Sub
  
catch:
  Err.Raise vbObjectError + 123, "XMLNodes.Add", "ID'ed node not unique"
End Sub


Public Function AddBy(ByVal eNodeType As XMLNodeTypes, Optional ByRef rstrName As String, Optional ByRef rstrText As String, Optional ByRef rstrID As String, Optional ByRef rstrValue As String) As XMLNode
  Dim xNode As XMLNode
  
  Select Case eNodeType
    Case xntNode, xntEmpty
      Set xNode = New XMLNode
      With xNode
        .NodeType = eNodeType
        .Name = rstrName
        If rstrText <> "" Then .Text = rstrText
        If rstrID <> "" Then .ID = rstrID
        If rstrValue <> "" Then .Value = rstrValue
        'If Not IsEmpty(vBoolValue) Then .BoolValue = vBoolValue
      End With
      Me.Add xNode
      
    Case xntComment, xntCDATA, xntText
      Set xNode = New XMLNode
      With xNode
        .NodeType = eNodeType
        .Text = rstrText
      End With
      Me.Add xNode
      
  End Select
  
  Set AddBy = xNode
End Function


Public Function Count() As Long
  Count = pcolNodes.Count
End Function


Public Sub Remove(ByVal vKey As Variant, Optional ByRef rstrID As String = "")
  If rstrID = "" Then
    pcolNodes.Remove vKey
  Else
    pcolNodes.Remove CStr(vKey) & " " & rstrID
  End If
End Sub


Public Sub Clear()
  Dim n As Long, m As Long
  n = pcolNodes.Count
  For m = n To 1 Step -1
    pcolNodes.Remove m
  Next
End Sub


Public Function IsElement(ByVal vKey As Variant, Optional ByRef rstrID As String = "") As Boolean
  Dim xNode As XMLNode
  On Error GoTo catch
  If rstrID = "" Then
    Set xNode = pcolNodes(vKey)  ' direct access by index
  Else
    Set xNode = pcolNodes(CStr(vKey) & " " & rstrID) ' vKey = node name
  End If
  IsElement = True
  Exit Function
catch:
End Function


Public Function Item(ByVal vKey As Variant, Optional ByRef rstrID As String = "") As XMLNode
Attribute Item.VB_UserMemId = 0
  If rstrID = "" Then
    Set Item = pcolNodes(vKey)  ' direct access by index
  Else
    Set Item = pcolNodes(CStr(vKey) & " " & rstrID) ' vKey = node name
  End If
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = pcolNodes.[_NewEnum]
End Function


Private Sub Class_Initialize()
  Set pcolNodes = New Collection
End Sub
