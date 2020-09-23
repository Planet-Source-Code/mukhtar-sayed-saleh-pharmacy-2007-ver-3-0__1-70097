VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frm_companies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÔÑßÇÊ ÇáÃÏæíÉ"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_companies.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "comname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "companies"
      RightToLeft     =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4440
      OleObjectBlob   =   "frm_companies.frx":29C12
      Top             =   3000
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   4815
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frm_companies.frx":29E46
         TabIndex        =   7
         Top             =   320
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   9015
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_companies.frx":29EB8
         Height          =   4935
         Left            =   0
         OleObjectBlob   =   "frm_companies.frx":29ECC
         TabIndex        =   10
         Top             =   0
         Width           =   9015
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "frm_companies.frx":2A70F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ ÔÑßÉ"
         Height          =   855
         Left            =   5760
         Picture         =   "frm_companies.frx":2AD0B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá ÔÑßÉ"
         Height          =   855
         Left            =   6840
         Picture         =   "frm_companies.frx":2B14D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ÔÑßÉ"
         Height          =   855
         Left            =   7920
         Picture         =   "frm_companies.frx":2B58F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_companies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇÓã ÇáÔÑßÉ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

Data1.Recordset.AddNew
Text2.Text = Text1.Text
On Error Resume Next
Data1.Recordset.MoveNext
Data1.Recordset.MovePrevious
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = ""
frm_main.Refreshcommand

End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÔÑßÉ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If
Data1.Recordset.Edit
Text2.Text = Text1.Text
On Error Resume Next
Data1.Recordset.MoveNext
Data1.Recordset.MovePrevious
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = ""
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÔÑßÉ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÔÑßÉ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = Text2.Text
frm_main.Refreshcommand

Else
Exit Sub
End If

End Sub


Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text2.Text <> "" Then
Text1.Text = Text2.Text
End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from companies"
Data1.Refresh

End Sub
