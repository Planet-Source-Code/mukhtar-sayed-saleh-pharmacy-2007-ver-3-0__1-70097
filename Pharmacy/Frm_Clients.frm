VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Clients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÓÌá ÇáÚãáÇÁ"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Clients.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "clientcode"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "clientname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Frm_Clients.frx":29C12
      Top             =   2760
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clients"
      RightToLeft     =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_Clients.frx":29E46
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "Frm_Clients.frx":29E5A
      TabIndex        =   5
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Clients.frx":2A82C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Úãíá ÌÏíÏ"
         Height          =   855
         Left            =   6360
         Picture         =   "Frm_Clients.frx":2AE28
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá Úãíá"
         Height          =   855
         Left            =   5280
         Picture         =   "Frm_Clients.frx":2B26A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ Úãíá"
         Height          =   855
         Left            =   4200
         Picture         =   "Frm_Clients.frx":2B6AC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_Clients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frm_AddClient.Show
StayOnTop Frm_AddClient
storecommand2 = "addnew"

End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
Frm_AddClient.Show
StayOnTop Frm_AddClient
storecommand2 = "edit"
Frm_AddClient.Text1.Text = Text1.Text
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÚãíá ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
End If

End Sub

Private Sub Command3_Click()
'íÌÈ Ãä íßæä ãÌãæÚ ãÏíä + ÏÇÆä = 0 ÍÊì íäÍÐÝ
With Frm_Clients_Money
.Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data2.RecordSource = "select * from clients_money where clientcode=" & Text2.Text & ""
.Data2.Refresh
.refreshall
If CDbl(.Text8.Text) = CDbl(.Text9.Text) Then
 Dim mok
 mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÇáÚãíá " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ")
 If mok = vbYes Then
  On Error Resume Next
  Data1.Recordset.Delete
  Data1.Refresh
  DBGrid1.ReBind
  DBGrid1.Refresh
     frm_main.Refreshcommand
 Else
  Exit Sub
 End If
Else
 MsgBox "áÇ íãßä ÍÐÝ ÇáÚãíá íÌÈ ÊÕÝíÉ ÇáÍÓÇÈ ÃæáÇð", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
 Exit Sub
End If
End With
Unload Frm_Clients_Money
End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from clients"
Data1.Refresh

End Sub

