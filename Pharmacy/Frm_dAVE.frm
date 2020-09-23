VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_dAVE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊæÝíÑ ÇáÏæÇÁ"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_dAVE.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "Frm_dAVE.frx":29C12
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton Command2 
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜÜæÇÝÜÜÜÜÞ"
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   6255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "Frm_dAVE.frx":29E46
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Frm_dAVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then
MsgBox "ÇáÑÌÇÁ ÇÏÎÇá ÑÞã ÕÍíÍ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
'ÌãÚ ÇáßãíøÉ ãÚ ÇáãÓÊæÏÚ
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Frm_Naf.Text1.Text & "' and doname='" & Frm_Naf.Text2.Text & "'"
.Data1.Refresh
.Data1.Recordset.Edit
.Text5.Text = CLng(.Text5.Text) + CLng(Text1.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
'ÍÐÝ ÇáßãíøÉ ãä ÞÇÆãÉ ÇáäæÇÝÐ
With Frm_Naf
.Data1.Recordset.Delete
.Data1.Refresh
.DBGrid1.ReBind
.DBGrid1.Refresh
End With
'ÊÍÏíË ÇáãÓÊæÏÚ
'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
With frm_store
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from pharstore"
.Data1.Refresh
.DBGrid1.ReBind
.DBGrid1.Refresh
End With
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
MsgBox "ÇÏÎá ÚÏÏ ÕÍíÍ ÝÞØ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text1.Text = ""
End If
End Sub
