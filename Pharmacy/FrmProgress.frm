VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÌÇÑí ãÚÇáÌÉ ÇáÈíÇäÇÊ"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar bar 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "FrmProgress.frx":0000
      Top             =   1320
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub
