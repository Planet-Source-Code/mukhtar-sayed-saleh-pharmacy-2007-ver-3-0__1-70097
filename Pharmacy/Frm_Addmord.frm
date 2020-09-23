VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_Addmord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÖÇÝÉ ãæÑøÏ ÌÏíÏ"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2880
         OleObjectBlob   =   "Frm_Addmord.frx":0000
         Top             =   1320
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "Frm_Addmord.frx":0234
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "Frm_Addmord.frx":02A6
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_Addmord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SkinLabel3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÅÏÎÇá ßÇÝÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If mordcommand1 = "addnew" Then
'ÇÖÇÝÊå Ýí ÌÏæÇá ÇáãæÑÏíä ÇæáÇð
With Frm_Mord
.Data1.Recordset.AddNew
.Text1.Text = CStr(Text1.Text)
.Text2.Text = CStr(Text2.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
'ÇÖÇÝÊå Ýí ÌÏæá ÇáÚãáÇÁ ËÇäíÇð
With Frm_Clients
.Data1.Recordset.AddNew
.Text1.Text = Text1.Text
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With

Unload Me
frm_main.Refreshcommand

End If

If mordcommand1 = "edit" Then
'ÊÚÏíá Ýí ÌÏæá ÇáÚãáÇÁ
With Frm_Clients
  .Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
  .Data1.RecordSource = "select * from clients where clientname='" & Frm_Mord.Text1.Text & "'"
  .Data1.Refresh
 
  If .Text1.Text <> "" Then
   .Data1.Recordset.Edit
   .Text1.Text = CStr(Text1.Text)
   On Error Resume Next
   .Data1.Recordset.MoveNext
   .Data1.Recordset.MovePrevious
  End If
End With

'ÊÚÏíá Ýí ÌÏæá ÇáãæÑÏíä
With Frm_Mord
.Data1.Recordset.Edit
.Text1.Text = CStr(Text1.Text)
.Text2.Text = CStr(Text2.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload Me
frm_main.Refreshcommand

End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub
