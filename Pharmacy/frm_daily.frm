VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_daily 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáíæãíøÉ"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_daily.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "daily"
      RightToLeft     =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5160
      OleObjectBlob   =   "frm_daily.frx":29C12
      Top             =   4320
   End
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   10575
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "date"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   2.45745e5
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "pamount"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2.45745e5
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_daily.frx":29E46
         Height          =   8055
         Left            =   0
         OleObjectBlob   =   "frm_daily.frx":29E5A
         TabIndex        =   6
         Top             =   120
         Width           =   10575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇÑÊÌÇÚ ÏæÇÁ"
         Height          =   855
         Left            =   6600
         Picture         =   "frm_daily.frx":2A824
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ãÈíÚ áÚãíá ÈÇáÃÌá"
         Height          =   855
         Left            =   7920
         Picture         =   "frm_daily.frx":2AE64
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "äÞØÉ ãÈíÚ"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_daily.frx":2B51F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "frm_daily.frx":2BC3C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_daily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_SalePoint.Show
StayOnTop frm_SalePoint
pointcommand = "cash"
End Sub

Private Sub Command2_Click()
frm_SalePoint.Show
StayOnTop frm_SalePoint
pointcommand = "client"

End Sub

Private Sub Command4_Click()
frm_change.Show
StayOnTop frm_change

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
'Data1.RecordSource = "select * from daily where date between #" & CDate(Format(Date, "Short Date")) & "# and #" & CDate(Format(Date, "Short Date")) & "#"
Data1.RecordSource = "select * from daily where date= #" & CDate(Format(Date, "Short Date")) & "#"
Data1.Refresh
DBGrid1.Caption = " íæãíøÉ " & Format(Date, "Short Date")
DBGrid1.Refresh
DBGrid1.ReBind

End Sub


