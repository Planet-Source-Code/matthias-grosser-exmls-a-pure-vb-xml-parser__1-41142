VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmXMLTest 
   Caption         =   "eXMLs"
   ClientHeight    =   4995
   ClientLeft      =   1635
   ClientTop       =   2400
   ClientWidth     =   10230
   Icon            =   "frmXMLTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   10230
   Begin RichTextLib.RichTextBox txtXML 
      Height          =   4515
      Left            =   225
      TabIndex        =   5
      Top             =   225
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   7964
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmXMLTest.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEnumDemo 
      Caption         =   "Enumerate Mammals"
      Height          =   390
      Left            =   8250
      TabIndex        =   4
      Top             =   2475
      Width           =   1815
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Mickey's Color"
      Height          =   390
      Left            =   8250
      TabIndex        =   3
      Top             =   2025
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdlLoad 
      Left            =   8625
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   390
      Left            =   8250
      TabIndex        =   2
      Top             =   1350
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadXML 
      Caption         =   "LoadXML"
      Height          =   390
      Left            =   8250
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   390
      Left            =   8250
      TabIndex        =   0
      Top             =   225
      Width           =   1215
   End
End
Attribute VB_Name = "frmXMLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public XD As XMLDocument

Private Sub cmdChange_Click()
  ' navigation inside the document is easy:
  XD.First.Children("universe", "thisone").Children("planet", "earth").First("lifeforms").Children("mammal", "mouse").Attributes("color") = InputBox("Change Mickey's color to:")
  txtXML.Text = XD.XML
End Sub

Private Sub cmdCreate_Click()
  Dim xn As XMLNode, xn2 As XMLNode, xn3 As XMLNode, xn4 As XMLNode, xn5 As XMLNode
  
  Set XD = New XMLDocument
  Set xn = XD.Children.AddBy(xntNode, "existance")
  xn.Attributes.Add "created", CStr(Now)
  Set xn2 = xn.Children.AddBy(xntNode, "universe", , "thisone")
  xn2.Attributes.Add "created", CStr(Now)
  With xn2.Children
    .AddBy xntNode, "cmd", "lights on!", "first"
    Set xn3 = .AddBy(xntNode, "planet", , "pluto")
    With xn3.Attributes
      .Add "sky", "black"
      .Add "weather", "bad"
      .Add "created", CStr(Now)
    End With
    With xn3.Children
      .AddBy xntComment, , "far too cold outta here"
      .AddBy xntEmpty, "lifeforms"
    End With
    Set xn3 = .AddBy(xntNode, "planet", , "earth")
  End With
  With xn3.Attributes
    .Add "sky", "always blue"
    .Add "weather", "fine"
    .Add "created", CStr(Now)
  End With
  Set xn4 = xn3.Children.AddBy(xntNode, "lifeforms")
  With xn4.Children
    .AddBy xntEmpty, "fish", , "dolphin", "flipper"
    .AddBy xntComment, , "you think dolphin isn't a fish? well... that's your problem!"
    .AddBy xntNode, "fish", "better don't touch me", "giant jellyfish"
    Set xn5 = .AddBy(xntEmpty, "plant", , "tree")
    With xn5.Attributes
      .Add "color", "green"
      .Add "intelligence", CStr(0)
    End With
    Set xn5 = .AddBy(xntEmpty, "plant", , "salad")
    With xn5.Attributes
      .Add "color", "green too"
      .Add "intelligence", CStr(-25)
    End With
    Set xn5 = .AddBy(xntEmpty, "mammal", , "elephant", "dumbo")
    With xn5.Attributes
      .Add "color", "grey"
      .Add "intelligence", CStr(5)
    End With
    Set xn5 = .AddBy(xntEmpty, "mammal", , "mouse", "mickey")
    With xn5.Attributes
      .Add "color"
      .Add "intelligence", CStr(20)
    End With
    Set xn5 = .AddBy(xntEmpty, "mammal", , "cow", "milka")
    With xn5.Attributes
      .Add "color", "lila"
      .Add "intelligence", CStr(-25000)
    End With
  End With
  Set xn2 = xn.Children.AddBy(xntNode, "universe", , "parallel")
  xn2.Children.AddBy xntComment, , " add any intelligent life forms here ;-) "
  Set xn2 = xn.Children.AddBy(xntNode, "universe", , "inverse")
  xn2.Children.AddBy xntNode, "lie", "1 + 1 = 0", "axiomatic", CXMLBool(True)
  xn2.Children.AddBy xntNode, "lie", "and that's the truth!", "another", CXMLBool(False)
  txtXML.Text = XD.XML
  
End Sub


Private Sub cmdEnumDemo_Click()
  Dim xn As XMLNode, strList As String
  For Each xn In XD.First.Children("universe", "thisone").Children("planet", "earth").First("lifeforms").EnumChildrenByName("mammal")
    strList = strList & "• " & xn.ID & ": " & xn.Value & vbLf
  Next
  MsgBox "Available mammals on planet earth: " & vbLf & vbLf & strList, vbInformation
End Sub

Private Sub cmdLoad_Click()
  cdlLoad.ShowOpen
  If cdlLoad.filename <> "" Then XD.Load cdlLoad.filename
  txtXML.Text = XD.XML
End Sub

Private Sub cmdLoadXML_Click()
  On Error GoTo catch
  XD.LoadXML txtXML.Text
  txtXML.Text = XD.XML
  Exit Sub
catch:
  MsgBox "Parse Error: " & Err.Source & ": " & Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_Initialize()
  Set XD = New XMLDocument
End Sub

Private Sub Form_Load()
  txtXML.Text = XD.XML
End Sub

Private Sub Form_Resize()
  txtXML.Width = Me.ScaleWidth - 2250
  txtXML.Height = Me.ScaleHeight - 450
  cmdCreate.Left = Me.ScaleWidth - 1890
  cmdLoad.Left = cmdCreate.Left
  cmdLoadXML.Left = cmdCreate.Left
  cmdChange.Left = cmdCreate.Left
  cmdEnumDemo.Left = cmdChange.Left
End Sub
