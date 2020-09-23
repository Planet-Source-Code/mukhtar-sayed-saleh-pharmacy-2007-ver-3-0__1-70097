VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Frm_Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÚáíãÇÊ ÇáäÙÇã"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Help.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4680
      OleObjectBlob   =   "Frm_Help.frx":29C12
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÅÛáÇÞ"
      Height          =   375
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   8280
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      ExtentX         =   20135
      ExtentY         =   14420
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Frm_Help"
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

WebBrowser1.Navigate App.Path & ("\help.mok")
End Sub

