VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_Increment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊãÏíÏ ÝÚÇáíÉ ÏæÇÁ"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "Frm_Increment.frx":0000
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÜæÇÝÜÞ"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Frm_Increment.frx":0234
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Frm_Increment.frx":02A6
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Frm_Increment.frx":0318
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Frm_Increment.frx":038A
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_Increment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text5.Text = "" Then
 MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÊÇÑíÎ ÇáÝÚøÇáíøÉ ÇáÌÏíÏ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
 Exit Sub
End If

If IsDate(Text5.Text) = False Then
 MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÊÇÑíÎ ÇáÝÚøÇáíøÉ ÇáÌÏíÏ ÈÔßá ÕÍíÍ áÃä ÇáÊÚÈíÑ ÇáãÏÎá áíÓ ÊÇÑíÎÇð", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
 Exit Sub
End If

'ÈÏÁ ÇáÚãáíøÉ
With frm_store
  .Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
  .Data1.RecordSource = "select * from pharstore where comname='" & Text1.Text & "' and doname='" & Text2.Text & "' and docode='" & Text3.Text & "'"
  .Data1.Refresh
 
  If .Text1.Text <> "" Then
   .Data1.Recordset.Edit
   .Text11.Text = CDate(Text5.Text)
   On Error Resume Next
   .Data1.Recordset.MoveNext
   .Data1.Recordset.MovePrevious
  End If

End With
MsgBox "ÊãÊ ÇáÚãáíÉ ÈäÌÇÍ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
Unload Me
frm_main.Initial
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub


Private Sub Text5_GotFocus()
Text5.Text = Format(Date + 110, "dd/mm/yyyy")
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)


End Sub
