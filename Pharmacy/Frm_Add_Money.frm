VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_Add_Money 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÖÇÝÉ ãÈáÛ"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Add_Money.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2760
      OleObjectBlob   =   "Frm_Add_Money.frx":29C12
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2640
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "Frm_Add_Money.frx":29E46
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1200
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "Frm_Add_Money.frx":29EAE
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "Frm_Add_Money.frx":29F14
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "Frm_Add_Money.frx":29F86
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "Frm_Add_Money.frx":29FEC
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_Add_Money"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÅÏÎÇá ßÇÝÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If IsNumeric(Text2.Text) = False Then
MsgBox "Åä ÇáÞíãÉ ÇáÊí Ýí ÍÞá ãÏíä áíÓÊ ÑÞãÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If IsNumeric(Text3.Text) = False Then
MsgBox "Åä ÇáÞíãÉ ÇáÊí Ýí ÍÞá ÏÇÆä áíÓÊ ÑÞãÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If IsDate(Text4.Text) = False Then
MsgBox "ÇáÊÚÈíÑ ÇáÐí Ýí ÍÞá ÇáÊÇÑíÎ áíÓ ÊÇÑíÎÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If Text2.Text = "0" And Text3.Text = "0" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇáÞíãÉ ÇáãÇáíøÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
If storecommand3 = "addnew" Then
With Frm_Clients_Money
.Data2.Recordset.AddNew
.Text3.Text = .Text2.Text
.Text4.Text = CDbl(Text2.Text)
.Text5.Text = CDbl(Text3.Text)
.Text6.Text = Format(Text4.Text, "Short Date")
.Text7.Text = Text5.Text
On Error Resume Next
.Data2.Recordset.MoveNext
.Data2.Recordset.MovePrevious
.DBGrid2.Refresh
.DBGrid2.ReBind
End With
Frm_Clients_Money.refreshall
Unload Me
End If

If storecommand3 = "edit" Then
With Frm_Clients_Money
.Data2.Recordset.Edit
.Text4.Text = CDbl(Text2.Text)
.Text5.Text = CDbl(Text3.Text)
.Text6.Text = Format(Text4.Text, "Short Date")
.Text7.Text = Text5.Text
On Error Resume Next
.Data2.Recordset.MoveNext
.Data2.Recordset.MovePrevious
.DBGrid2.Refresh
.DBGrid2.ReBind
End With
Frm_Clients_Money.refreshall
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


Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Text3.Text = "0"
End Sub


Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
Text2.Text = "0"

End Sub
