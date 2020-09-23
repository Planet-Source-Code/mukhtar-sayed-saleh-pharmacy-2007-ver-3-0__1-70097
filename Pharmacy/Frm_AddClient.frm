VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_AddClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÖÇÝÉ Úãíá ÌÏíÏ"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_AddClient.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "Frm_AddClient.frx":29C12
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜÜæÇÝÜÜÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "Frm_AddClient.frx":29E46
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Frm_AddClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇÓã ÇáÚãíá", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
If storecommand2 = "addnew" Then
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

If storecommand2 = "edit" Then
With Frm_Clients
.Data1.Recordset.Edit
.Text1.Text = Text1.Text
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload Me

End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub
