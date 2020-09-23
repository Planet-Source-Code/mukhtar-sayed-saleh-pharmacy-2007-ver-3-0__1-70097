VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Accounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÍÓÇÈÇÊ ÇáãÓÊÎÏãíä"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   ControlBox      =   0   'False
   Icon            =   "Frm_Accounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Accounts.frx":1ABEA
   RightToLeft     =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      DataField       =   "shapesedit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "Text11"
      Top             =   65464
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "mordenedit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text10"
      Top             =   65464
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      DataField       =   "Settings"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text9"
      Top             =   2.45745e5
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      DataField       =   "Reports"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text8"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "Daily"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   2.45745e5
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "clientedit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "storeedit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "companyedit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "type"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "userpass"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "username"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Frm_Accounts.frx":447FC
      Top             =   2400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      RightToLeft     =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ ãÓÊÎÏã"
         Height          =   855
         Left            =   5040
         Picture         =   "Frm_Accounts.frx":44A30
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Accounts.frx":44E72
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá ãÓÊÎÏã"
         Height          =   855
         Left            =   6120
         Picture         =   "Frm_Accounts.frx":4546E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ãÓÊÎÏã ÌÏíÏ"
         Height          =   855
         Left            =   7200
         Picture         =   "Frm_Accounts.frx":45BCB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_Accounts.frx":4633A
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "Frm_Accounts.frx":4634E
      TabIndex        =   16
      Top             =   1080
      Width           =   8295
   End
End
Attribute VB_Name = "Frm_Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
With Frm_NewAccount
  .Text1.Text = Text1.Text
  .Text2.Text = Text2.Text
  If CBool(Text3.Text) = True Then
  .Check1.Value = 1
  Else
  .Check1.Value = 0
  End If
  If CBool(Text4.Text) = True Then
  .Check2.Value = 1
  Else
  .Check2.Value = 0
  End If
    If CBool(Text5.Text) = True Then
  .Check3.Value = 1
  Else
  .Check3.Value = 0
  End If
  If CBool(Text6.Text) = True Then
  .Check4.Value = 1
  Else
  .Check4.Value = 0
  End If
  If CBool(Text7.Text) = True Then
  .Check5.Value = 1
  Else
  .Check5.Value = 0
  End If
  If CBool(Text8.Text) = True Then
  .Check6.Value = 1
  Else
  .Check6.Value = 0
  End If
  If CBool(Text9.Text) = True Then
  .Check7.Value = 1
  Else
  .Check7.Value = 0
  End If

End With

Frm_NewAccount.Show
StayOnTop Frm_NewAccount
mokcommand = "edit"

End Sub

Private Sub Command2_Click()
Frm_NewAccount.Show
StayOnTop Frm_NewAccount
mokcommand = "new"
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáãÓÊÎÏã", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If
 

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÍÓÇÈ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.refreshcommands
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from users"
Data1.Refresh
Me.refreshcommands
End Sub

Public Function refreshcommands()
If Text1.Text = nowuser Then
Command1.Enabled = False
Command3.Enabled = False
Else
Command1.Enabled = True
Command3.Enabled = True
End If
End Function
