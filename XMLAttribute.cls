VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private variables =========================================================================================

Private pstrID As String
Private pstrValue As String


' public properties =========================================================================================

Public Property Let ID(ByRef rstrID As String)
  If pstrID <> "" Or rstrID = "" Then Err.Raise vbObjectError, "XMLAttribute.ID", "ID empty or already set"
  pstrID = rstrID
End Property

Public Property Get ID() As String
  ID = pstrID
End Property


Public Property Let Value(ByRef rstrValue As String)
  If StrComp(pstrID, "ID", vbTextCompare) = 0 And pstrValue <> "" Then Err.Raise vbObjectError, "XMLAttribute.Value", "Changing ID attribute value not allowed"
  pstrValue = rstrValue
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
  Value = pstrValue
End Property


Public Property Get XML() As String
  XML = pstrID & "=""" & UTF8Encode(XMLEntityEncode(Value)) & """"
End Property
