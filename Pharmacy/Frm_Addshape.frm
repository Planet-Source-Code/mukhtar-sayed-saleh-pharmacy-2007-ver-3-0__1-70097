VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_Addshape 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ôßá ÚÈæÉ ÌÏíÏ"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1155
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2880
         OleObjectBlob   =   "Frm_Addshape.frx":0000
         Top             =   1320
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_Addshape.frx":0234
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frm_Addshape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÅÏÎÇá ßÇÝÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If shapecommand1 = "addnew" Then
With Frm_Shapes
.Data1.Recordset.AddNew
.Text1.Text = CStr(Text1.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload Me
frm_main.Refreshcommand

End If

If shapecommand1 = "edit" Then
With Frm_Shapes
.Data1.Recordset.Edit
.Text1.Text = CStr(Text1.Text)
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
