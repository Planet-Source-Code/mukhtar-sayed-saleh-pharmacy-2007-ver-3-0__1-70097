VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Íæá ÇáäÙÇã"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ControlBox      =   0   'False
   Icon            =   "frm_about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_about.frx":1ABEA
   RightToLeft     =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3840
      OleObjectBlob   =   "frm_about.frx":447FC
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Height          =   495
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5160
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5200
      Left            =   240
      Picture         =   "frm_about.frx":44A30
      RightToLeft     =   -1  'True
      ScaleHeight     =   5205
      ScaleWidth      =   9405
      TabIndex        =   0
      Top             =   0
      Width           =   9400
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub
