VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_NewAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅäÔÇÁ ÍÓÇÈ ãÓÊÎÏã ÌÏíÏ"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_NewAccount.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1800
      OleObjectBlob   =   "Frm_NewAccount.frx":29C12
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÜÜæÇÝÜÜÜÞ"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1080
      Width           =   4695
      Begin VB.CheckBox Check9 
         Caption         =   "                ÇáæÕæá ááãæÑÏíä"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check8 
         Caption         =   "        ÇáæÕæá áÃÔßÇá ÇáÚÈæÉ"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check7 
         Caption         =   "               ÇáæÕæá ááÅÚÏÇÏÇÊ"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         Caption         =   "                  ÇáæÕæá ááÊÞÇÑíÑ"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check5 
         Caption         =   "                   ÇáæÕæá ááíæãíøÉ"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Caption         =   "                   ÇáæÕæá ááÚãáÇÁ"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "               ÇáæÕæá ááãÓÊæÏÚ"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "      ÇáæÕæá áÔÑßÇÊ ÇáÃÏæíÉ "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "                         ãÏíÑ ááäÙÇã"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frm_NewAccount.frx":29E46
         TabIndex        =   17
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "Frm_NewAccount.frx":29EC4
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "Frm_NewAccount.frx":29F38
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_NewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÃßøÏ ãä ÊÚÈÆÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If mokcommand = "new" Then
  With Frm_Accounts
  .Data1.Recordset.AddNew
  .Text1.Text = Text1.Text
  .Text2.Text = Text2.Text
  .Text3.Text = CBool(Check1.Value)
  .Text4.Text = CBool(Check2.Value)
  .Text5.Text = CBool(Check3.Value)
  .Text6.Text = CBool(Check4.Value)
  .Text7.Text = CBool(Check5.Value)
  .Text8.Text = CBool(Check6.Value)
  .Text9.Text = CBool(Check7.Value)
  .Text10.Text = CBool(Check9.Value)
  .Text11.Text = CBool(Check8.Value)
  On Error Resume Next
  .Data1.Recordset.MoveNext
  .Data1.Recordset.MovePrevious
  .DBGrid1.Refresh
  .DBGrid1.ReBind
 End With
 Unload Me
End If

If mokcommand = "edit" Then
  With Frm_Accounts
  .Data1.Recordset.Edit
  .Text1.Text = Text1.Text
  .Text2.Text = Text2.Text
  .Text3.Text = CBool(Check1.Value)
  .Text4.Text = CBool(Check2.Value)
  .Text5.Text = CBool(Check3.Value)
  .Text6.Text = CBool(Check4.Value)
  .Text7.Text = CBool(Check5.Value)
  .Text8.Text = CBool(Check6.Value)
  .Text9.Text = CBool(Check7.Value)
  .Text10.Text = CBool(Check9.Value)
  .Text11.Text = CBool(Check8.Value)
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

