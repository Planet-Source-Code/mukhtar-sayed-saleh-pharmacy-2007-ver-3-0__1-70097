VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Mord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÞÇÆãÉ ÇáãæÑøÏíä"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ãæÑøÏ"
         Height          =   975
         Left            =   4920
         Picture         =   "Frm_Mord.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá ãæÑøÏ"
         Height          =   975
         Left            =   3720
         Picture         =   "Frm_Mord.frx":083E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ ãæÑøÏ"
         Height          =   975
         Left            =   2520
         Picture         =   "Frm_Mord.frx":1041
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   975
         Left            =   0
         Picture         =   "Frm_Mord.frx":1873
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "phone"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   64654
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "morname"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   64654
         Width           =   855
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Frm_Mord.frx":1E6F
         Height          =   4815
         Left            =   0
         OleObjectBlob   =   "Frm_Mord.frx":1E83
         TabIndex        =   1
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "morden"
      RightToLeft     =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4440
      OleObjectBlob   =   "Frm_Mord.frx":284D
      Top             =   3120
   End
End
Attribute VB_Name = "Frm_Mord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frm_Addmord.Show
Frm_Addmord.Caption = "ÅÖÇÝÉ ãæÑøÏ ÌÏíÏ"
StayOnTop Frm_Addmord
mordcommand1 = "addnew"

End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
Frm_Addmord.Show
StayOnTop Frm_Addmord
mordcommand1 = "edit"
Frm_Addmord.Text1.Text = Text1.Text
Frm_Addmord.Text2.Text = Text2.Text
Frm_Addmord.Caption = "ÊÚÏíá ãæÑøÏ"
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáãæÑøÏ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
End If

End Sub

Private Sub Command3_Click()


If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáãæÑøÏ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÇáãæÑøÏ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Else
Exit Sub
End If
frm_main.Refreshcommand

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from morden"
Data1.Refresh

End Sub
