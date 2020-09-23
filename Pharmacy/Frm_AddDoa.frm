VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_AddDoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÖÇÝÉ ÏæÇÁ ÌÏíÏ"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_AddDoa.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command6 
         Caption         =   "ÊæáíÏ ßæÏ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3960
         Width           =   4335
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   4440
         Width           =   4335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   3975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   3000
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   4920
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2880
         OleObjectBlob   =   "Frm_AddDoa.frx":29C12
         Top             =   1320
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":29E46
         TabIndex        =   14
         Top             =   3600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":29EC8
         TabIndex        =   15
         Top             =   2640
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":29F48
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":29FBA
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A02C
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A09E
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A120
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A192
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A206
         TabIndex        =   24
         Top             =   4560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "Frm_AddDoa.frx":2A27C
         TabIndex        =   25
         Top             =   4080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frm_AddDoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÅÏÎÇá ßÇÝÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If IsNumeric(Text4.Text) = False Then
MsgBox "ÓÚÑ ÇáÞØÚÉ ÇáæÇÍÏÉ ááÚãæã íÌÈ Ãä íßæä ÑÞãÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If IsNumeric(Text6.Text) = False Then
MsgBox "ÓÚÑ ÇáÞØÚÉ ÇáæÇÍÏÉ ááÕíÏáí íÌÈ Ãä íßæä ÑÞãÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If IsNumeric(Text5.Text) = False Then
MsgBox "ÚÏÏ ÇáÞØÚ ÇáãÊæÝøÑÉ íÌÈ Ãä íßæä ÑÞãÇð", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If IsDate(Text10.Text) = False Then
MsgBox "Åä ÇáÊÚÈíÑ ÇáãæÌæÏ Ýí ÍÞá ÊÇÑíÎ ÇáÝÚÇáíøÉ áíÓ ÊÇÑíÎÇð ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÇáÃãÑ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

If CDbl(Text6.Text) > CDbl(Text4.Text) Then
MsgBox "ÓÚÑ ÇáÏæÇÁ ááÚãæã íÌÈ Ãä íßæä ÃßÈÑ ãä ÓÚÑå ááÕíÏáí", vbInformation, "Êã ÅíÞÇÝ ÊÓÌíá ÇáÏæÇÁ"
Exit Sub
End If

If storecommand1 = "addnew" Then

'ÇáÊÍÞÞ ãä ÚÏã ÊÔÇÈå ßæÏ ÇáÏæÇÁ ãÚ ÏæÇÁ ËÇäí
  With frm_store.Data1.Recordset
  On Error Resume Next
      .MoveFirst
      .FindFirst "[docode] like '" & CStr(Text3.Text) & "*'"
    If .NoMatch = False Then
      MsgBox "íæÌÏ ÏæÇÁ ÂÎÑ áå äÝÓ ÇáßæÏ ÇáÑÌÇÁ ÊÛííÑ ßæÏ ÇáÏæÇÁ æ ÇáãÍÇæáÉ áÇÍÞÇð", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
      Exit Sub
    End If
  End With

With frm_store
.Data1.Recordset.AddNew
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
.Text3.Text = Text3.Text
.Text4.Text = CDbl(Text4.Text)
.Text5.Text = CLng(Text5.Text)
.Text6.Text = Text7.Text
.Text7.Text = Text8.Text
.Text8.Text = CDbl(Text6.Text)
.Text10.Text = Text9.Text
.Text11.Text = Format(CDate(Text10), "dd/mm/yyyy")
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload Me
End If

If storecommand1 = "edit" Then
With frm_store
.Data1.Recordset.Edit
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
.Text3.Text = Text3.Text
.Text4.Text = CDbl(Text4.Text)
.Text5.Text = CLng(Text5.Text)
.Text6.Text = Text7.Text
.Text7.Text = Text8.Text
.Text8.Text = CDbl(Text6.Text)
.Text10.Text = Text9.Text
.Text11.Text = Format(CDate(Text10), "dd/mm/yyyy")

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

Private Sub Command3_Click()
Dim I As Integer
frm_list.List1.Clear
With frm_companies
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from companies"
.Data1.Refresh
On Error Resume Next
For I = 0 To .Data1.Recordset.RecordCount - 1
frm_list.List1.AddItem .Text2.Text
.Data1.Recordset.MoveNext
Next
End With
frm_list.Caption = "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÔÑßÉ ÇáãØáæÈÉ"
frm_list.Show
StayOnTop frm_list

End Sub

Private Sub Command4_Click()
Frm_List6.Show
StayOnTop Frm_List6

End Sub

Private Sub Command5_Click()
Frm_List7.Show
StayOnTop Frm_List7

End Sub

Private Sub Command6_Click()
Dim mok As Boolean
mok = False
Do While (mok = False)
  bcodegen (CLng(frm_store.Data1.Recordset.RecordCount) + 1)
  With frm_store.Data1.Recordset
  On Error Resume Next
      .MoveFirst
      .FindFirst "[docode] like '" & str2 & "*'"
    If .NoMatch = True Then
     Text3.Text = bcodenew
     mok = True
    End If
  End With
Loop
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Text1_DblClick()
Dim I As Integer
frm_list.List1.Clear
With frm_companies
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from companies"
.Data1.Refresh
On Error Resume Next
For I = 0 To .Data1.Recordset.RecordCount - 1
frm_list.List1.AddItem .Text2.Text
.Data1.Recordset.MoveNext
Next
End With
frm_list.Caption = "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÔÑßÉ ÇáãØáæÈÉ"
frm_list.Show
StayOnTop frm_list

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim I As Integer
frm_list.List1.Clear
With frm_companies
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from companies"
.Data1.Refresh
On Error Resume Next
For I = 0 To .Data1.Recordset.RecordCount - 1
frm_list.List1.AddItem .Text2.Text
.Data1.Recordset.MoveNext
Next
End With
frm_list.Caption = "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÔÑßÉ ÇáãØáæÈÉ"
frm_list.Show
StayOnTop frm_list
End If

End Sub


Private Sub Text10_GotFocus()
Text10.Text = Format(Date + 210, "dd/mm/yyyy")
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If

End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If

End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If

End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub


Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text4.Text)

End Sub

Private Sub Text7_DblClick()
Frm_List6.Show
StayOnTop Frm_List6

End Sub


Private Sub Text8_DblClick()
Frm_List7.Show
StayOnTop Frm_List7

End Sub
