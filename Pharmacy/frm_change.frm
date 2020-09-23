VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_change 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇÑÊÌÇÚ ÏæÇÁ"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_change.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3960
      OleObjectBlob   =   "frm_change.frx":29C12
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   8055
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   405
            Left            =   5160
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   330
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   405
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   250
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   330
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            MaxLength       =   4
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Text            =   "0"
            Top             =   360
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   5160
            OleObjectBlob   =   "frm_change.frx":29E46
            TabIndex        =   7
            Top             =   120
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   2400
            OleObjectBlob   =   "frm_change.frx":29EB8
            TabIndex        =   8
            Top             =   120
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frm_change.frx":29F2A
            TabIndex        =   9
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÜÞ"
         Default         =   -1  'True
         Height          =   495
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frm_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇÎÊíÇÑ ÇáÏæÇÁ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If Text3.Text = "" Or Text3.Text = CLng(0) Then
MsgBox "ÇáÑÌÇÁ ÇÏÎÇá ÚÏÏ ÇáÞØÚ ÇáãÑÊÌÚÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If IsNumeric(Text3.Text) = False Then
MsgBox "ÇÏÎá ÇÑÞÇã ÕÍíÍÉ ÝÞØ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text3.Text = "0"
Exit Sub
End If

'ÌãÚ ÇáßãíøÉ ãÚ ÇáãÓÊæÏÚ
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Text1.Text & "' and doname='" & Text2.Text & "'"
.Data1.Refresh
.Data1.Recordset.Edit
.Text5.Text = CLng(.Text5.Text) + CLng(Text3.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload frm_store
'ÇÖÇÝÉ ÇáÓÌá ááíæãíÉ
With frm_daily
.Data1.Recordset.AddNew
.Text1.Text = CDbl(-1) * (CDbl(CLng(Text3.Text)) * CDbl(selcash2))
.Text2.Text = CDate(Format(Date, "Short Date"))
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
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

Private Sub Text1_DblClick()
frm_list4.Show
StayOnTop frm_list4

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_list4.Show
StayOnTop frm_list4
Else
Exit Sub
End If

End Sub

Private Sub Text2_DblClick()
frm_list4.Show
StayOnTop frm_list4

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_list4.Show
StayOnTop frm_list4
Else
Exit Sub
End If

End Sub

Private Sub Text3_Change()
If Text3.Text = "" Or Text3.Text = "0" Then
Exit Sub
Else
If IsNumeric(Text3.Text) = False Then
MsgBox "ÇÏÎá ÇÑÞÇã ÕÍíÍÉ ÝÞØ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text3.Text = "0"
Exit Sub
End If
End If

End Sub
