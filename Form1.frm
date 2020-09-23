VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin Project1.DSkin DSkin1 
      Left            =   2280
      Top             =   2280
      _ExtentX        =   3836
      _ExtentY        =   1296
      Visual_Refresh_Interval=   "10000"
      Skin            =   "C:\My Documents\Vb Stuff\Browser\skins\green.skn"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
