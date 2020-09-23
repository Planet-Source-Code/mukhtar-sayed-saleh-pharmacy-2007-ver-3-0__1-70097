VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_store_find 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÈÍË æ ÊÕÝíÉ"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_store_find.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3120
      OleObjectBlob   =   "frm_store_find.frx":29C12
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   2175
         Begin VB.OptionButton Option3 
            Caption         =   "                    ßæÏ ÇáÏæÇÁ"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "                     ÇÓã ÇáÏæÇÁ"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "                   ÇÓã ÇáÔÑßÉ"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frm_store_find.frx":29E46
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãæÇÝÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_store_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ãÚíÇÑ ÇáÊÕÝíÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If Option1.Value = True Then
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Text1.Text & "'"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
 If .Text1.Text <> "" Then
 Unload Me
 Else
 MsgBox "áÇ íæÌÏ äÊÇÆÌ ãØÇÈÞÉ áÈÍËß", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
 .Data1.RecordSource = "select * from pharstore"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
 Exit Sub
 End If
End With
End If

If Option2.Value = True Then
With frm_store
.Data1.RecordSource = "select * from pharstore where doname='" & Text1.Text & "'"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
 If .Text1.Text <> "" Then
 Unload Me
 Else
 MsgBox "áÇ íæÌÏ äÊÇÆÌ ãØÇÈÞÉ áÈÍËß", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
  .Data1.RecordSource = "select * from pharstore"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
 Exit Sub
 End If
End With

End If

If Option3.Value = True Then
With frm_store
.Data1.RecordSource = "select * from pharstore where docode='" & Text1.Text & "'"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
 If .Text1.Text <> "" Then
 Unload Me
 Else
 MsgBox "áÇ íæÌÏ äÊÇÆÌ ãØÇÈÞÉ áÈÍËß", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
  .Data1.RecordSource = "select * from pharstore"
.Data1.Refresh
.DBGrid1.Refresh
.DBGrid1.ReBind
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
