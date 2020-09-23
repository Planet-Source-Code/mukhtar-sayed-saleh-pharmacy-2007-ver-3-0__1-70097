VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÑÌÇÁ ÇáÊÍÏíÏ"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_list.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2280
      OleObjectBlob   =   "frm_list.frx":29C12
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frm_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
selected = List1.Text
Frm_AddDoa.Text1.Text = selected
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hwnd
End Sub

Private Sub List1_Click()
selected = List1.Text
End Sub
