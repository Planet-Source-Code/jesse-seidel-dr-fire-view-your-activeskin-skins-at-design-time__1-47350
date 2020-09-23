VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.UserControl DSkin 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "10000"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   1560
   End
   Begin ACTIVESKINLibCtl.Skin skn1 
      Left            =   2280
      OleObjectBlob   =   "DSkin.ctx":0000
      Top             =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ActiveSkin Design-Time Fix By SpitFire"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "DSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
Timer1.Interval = Text1.Text
skn1.LoadSkin Text2.Text
SkinHWND 0
End Sub

Private Sub UserControl_Initialize()
Text2.Text = App.Path & "\skins\green.skn"
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Shape1.Height
UserControl.Width = Shape1.Width
End Sub

Public Function SkinHWND(hwnd As Long)
skn1.ApplySkin hwnd
End Function

Public Property Get Visual_Refresh_Interval() As String
Attribute Visual_Refresh_Interval.VB_Description = "Returns/sets the text contained in the control."
    Visual_Refresh_Interval = Text1.Text
End Property

Public Property Let Visual_Refresh_Interval(ByVal New_Visual_Refresh_Interval As String)
    Text1.Text() = New_Visual_Refresh_Interval
    PropertyChanged "Visual_Refresh_Interval"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Text1.Text = PropBag.ReadProperty("Visual_Refresh_Interval", "Text1")
    Text2.Text = PropBag.ReadProperty("Skin", "Text2")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Visual_Refresh_Interval", Text1.Text, "Text1")
    Call PropBag.WriteProperty("Skin", Text2.Text, "Text2")
End Sub

Public Property Get Skin() As String
Attribute Skin.VB_Description = "Returns/sets the text contained in the control."
    Skin = Text2.Text
End Property

Public Property Let Skin(ByVal New_Skin As String)
    Text2.Text() = New_Skin
    PropertyChanged "Skin"
End Property

