VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_Clients_Find 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÈÍË æ ÊÕÝíÉ"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Cliens_Find.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "Frm_Cliens_Find.frx":29C12
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÛíÑ ãÊÃßøÏ ãä ÇáÅÓã"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜæÇÝÜÜÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "Frm_Cliens_Find.frx":29E46
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Frm_Clients_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇÓã ÇáÚãíá", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
On Error Resume Next
If Check1.Value = 0 Then
With Frm_Clients_Money
.Data1.RecordSource = "select * from clients where clientname='" & Text1.Text & "'"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
.Data2.RecordSource = "select * from clients_money where clientcode=" & .Text2.Text & ""
.Data2.Refresh
.DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & .Text1.Text
.DBGrid2.Refresh
.DBGrid2.ReBind
.refreshall
 If .Text1.Text <> "" Then
 Unload Me
 Else
 MsgBox "áÇ íæÌÏ äÊÇÆÌ ãØÇÈÞÉ áÈÍËß", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
 .Data1.RecordSource = "select * from clients"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
.Data2.RecordSource = "select * from clients_money where clientcode=" & .Text2.Text & ""
.Data2.Refresh
.DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & .Text1.Text
.DBGrid2.Refresh
.DBGrid2.ReBind
.refreshall
 Exit Sub
 End If
End With
End If

If Check1.Value = 1 Then
With Frm_Clients_Money
.Data1.RecordSource = "select * from clients where clientname like """ & "*" & Text1.Text & "*" & """"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
.Data2.RecordSource = "select * from clients_money where clientcode=" & .Text2.Text & ""
.Data2.Refresh
.DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & .Text1.Text
.DBGrid2.Refresh
.DBGrid2.ReBind
.refreshall
 If .Text1.Text <> "" Then
 Unload Me
 Else
 MsgBox "áÇ íæÌÏ äÊÇÆÌ ãØÇÈÞÉ áÈÍËß", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
 .Data1.RecordSource = "select * from clients"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
.Data2.RecordSource = "select * from clients_money where clientcode=" & .Text2.Text & ""
.Data2.Refresh
.DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & .Text1.Text
.DBGrid2.Refresh
.DBGrid2.ReBind
.refreshall
 Exit Sub
 End If
End With
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub
