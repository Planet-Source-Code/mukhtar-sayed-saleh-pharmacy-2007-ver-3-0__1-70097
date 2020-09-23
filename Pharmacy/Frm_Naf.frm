VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Naf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÃÏæíÉ ÇáäÇÝÐÉ"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Naf.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "docount"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   65454
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4800
      OleObjectBlob   =   "Frm_Naf.frx":29C12
      Top             =   2760
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "peiceprice"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2.45745e5
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "docode"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "doname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "naf"
      RightToLeft     =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊæÝíÑ ÇáÏæÇÁ"
         Height          =   855
         Left            =   8760
         Picture         =   "Frm_Naf.frx":29E46
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Naf.frx":2A4EF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_Naf.frx":2AAEB
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "Frm_Naf.frx":2AAFF
      TabIndex        =   0
      Top             =   1080
      Width           =   9855
   End
End
Attribute VB_Name = "Frm_Naf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" Then
Frm_dAVE.Show
StayOnTop Frm_dAVE
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ ÇáÐí ÊÑÛÈ Ýí ÊæÝíÑ ßãíøÉ ãäå", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hwnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from naf"
Data1.Refresh

End Sub
