VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frm_SalePoint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "äÞØÉ ãÈíÚ"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_SalePoint.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      DataField       =   "dprice"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Text            =   "Text8"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "damount"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Text            =   "Text7"
      Top             =   2.45745e5
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "dname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "compn"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "temp"
      RightToLeft     =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "frm_SalePoint.frx":29C12
      Top             =   4200
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6360
      Width           =   10455
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ãÈíÚ ÌÏíÏ"
         Height          =   855
         Left            =   9240
         Picture         =   "frm_SalePoint.frx":29E46
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   120
         Picture         =   "frm_SalePoint.frx":2A563
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ÇáßãíøÉ"
         Height          =   855
         Left            =   6600
         Picture         =   "frm_SalePoint.frx":2AB5F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇÎÊíÇÑ ÇáÏæÇÁ"
         Height          =   855
         Left            =   7800
         Picture         =   "frm_SalePoint.frx":2AC28
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅäåÇÁ ÇáãÈíÚ"
         Height          =   855
         Left            =   4200
         Picture         =   "frm_SalePoint.frx":2B34C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅäÞÇÕ ÇáßãíøÉ"
         Height          =   855
         Left            =   5400
         Picture         =   "frm_SalePoint.frx":2B3F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_SalePoint.frx":2B47D
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frm_SalePoint.frx":2B491
      TabIndex        =   8
      Top             =   3480
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   10455
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   330
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "frm_SalePoint.frx":2C1D0
         TabIndex        =   13
         Top             =   120
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "frm_SalePoint.frx":2C242
         TabIndex        =   14
         Top             =   120
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frm_SalePoint.frx":2C2B4
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frm_SalePoint.frx":2C334
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.TextBox Text_sum 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Ñ.Ó.þ"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   60
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1785
         Left            =   0
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   600
         Width           =   10455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   495
         Left            =   8880
         OleObjectBlob   =   "frm_SalePoint.frx":2C39C
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_SalePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇáÊÃßÏ ãä ÊÚÈÆÉ ßÇÝÉ ÇáÍÞæá ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If CInt(Text3.Text) < CInt(1) Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÚÏÏ ÇáÞØÚ ÇáãÈÇÚÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

'ØÑÍ ÇáßãíøÉ ãä ÇáãÓÊæÏÚ
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Text1.Text & "' and doname='" & Text2.Text & "'"
.Data1.Refresh
.Data1.Recordset.Edit
.Text5.Text = CLng(.Text5.Text) - CLng(Text3.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload frm_store
'ÇÖÇÝÉ ááÞÇÆãÉ
Data1.Recordset.AddNew
Text5.Text = Text1.Text
Text6.Text = Text2.Text
Text7.Text = CInt(Text3.Text)
Text8.Text = CDbl(Text4.Text)
On Error Resume Next
Data1.Recordset.MoveNext
Data1.Recordset.MovePrevious
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = ""
Me.refreshall


End Sub

Private Sub Command2_Click()
frm_list2.Show
StayOnTop frm_list2

End Sub

Private Sub Command3_Click()
If pointcommand = "cash" Then
'ÇäåÇÁ ÇáãÈíÚ ÈÍÐÝ ÇáßãíÉ ãä åæä æ ÇÖÇÝÊåÇ ááíæãíÉ
If Text_sum.Text = "" Or Text_sum.Text = CDbl(0) Then
Exit Sub
End If

frm_saleend.Show
StayOnTop frm_saleend
frm_saleend.Text_sum = Text_sum.Text
End If

If pointcommand = "client" Then
summ1 = CDbl(Text_sum.Text)
Frm_EndClient.Show
StayOnTop Frm_EndClient
End If
End Sub

Private Sub Command4_Click()
If Text5.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ ÇáãÑÇÏ ÅäÞÇÕå", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÅäÞÇÕ ÇáßãíøÉ ¿ ", 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ")
If mok = vbYes Then
'ÌãÚ ÇáßãíøÉ ãÚ ÇáãÓÊæÏÚ
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Text5.Text & "' and doname='" & Text6.Text & "'"
.Data1.Refresh
.Data1.Recordset.Edit
.Text5.Text = CLng(.Text5.Text) + CLng(Text7.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload frm_store

'ÇäÞÇÕ ÇáÞÇÆãÉ
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = ""
Me.refreshall
Else
Exit Sub
End If

End Sub

Private Sub Command5_Click()
Dim mok
mok = MsgBox("åá ÃäÊ ãÊÃßøÏ ¿", 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
Dim I As Integer
On Error Resume Next
Data1.Recordset.MoveFirst
For I = 1 To Data1.Recordset.RecordCount
'ÇÖÇÝÉ ÇáßãíøÉ ááãÓÊæÏÚ ÞÈá ÍÐÝ ÇáÓÌá
With frm_store
.Data1.RecordSource = "select * from pharstore where comname='" & Text5.Text & "' and doname='" & Text6.Text & "'"
.Data1.Refresh
.Data1.Recordset.Edit
.Text5.Text = CLng(.Text5.Text) + CLng(Text7.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Unload frm_store

'ÍÐÝ ÇáÓÌá
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
Next
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = ""
Me.refreshall

Else
Exit Sub
End If
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from temp"
Data1.Refresh
Me.refreshall
End Sub


Private Sub SkinLabel8_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_DblClick()
frm_list2.Show
StayOnTop frm_list2

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_list2.Show
StayOnTop frm_list2
Else
Exit Sub
End If
End Sub

Private Sub Text2_DblClick()
frm_list2.Show
StayOnTop frm_list2

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_list2.Show
StayOnTop frm_list2
Else
Exit Sub
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Or Text3.Text = "0" Then
Exit Sub
Else
If IsNumeric(Text3.Text) = True Then
 If CInt(Text3.Text) > CInt(selcount) Then
 MsgBox "ÚÏÏ ÇáÞØÚ ÇáãÊæÝÑÉ " & selcount & " ÝÞØ ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
 Text3.Text = "0"
 End If
Text4.Text = (CDbl(CInt(Text3.Text)) * CDbl(selcash))
Me.refreshall
Else
MsgBox "ÇÏÎá ÇÑÞÇã ÕÍíÍÉ ÝÞØ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text3.Text = "0"
Exit Sub
End If
End If
End Sub

Public Function refreshall()
'ÓÚÑ ÇáãÈÇÚ ÍÇáíÇð
If IsNumeric(Text3.Text) Then
Text4.Text = (CDbl(CInt(Text3.Text)) * CDbl(selcash))
Else
Exit Function
End If
'ÊÍÏíË ÇáãÌãæÚ
Dim I As Integer
On Error Resume Next
Data1.Recordset.MoveFirst
Text_sum.Text = CDbl(0)
For I = 1 To Data1.Recordset.RecordCount
Text_sum.Text = CDbl(Text_sum.Text) + CDbl(Text8.Text)
Data1.Recordset.MoveNext
Next
'ÊÝÚíá ÒÑ ÇáÒÇÆÏ
If Text4.Text <> "" And Text4.Text <> CDbl(0) Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
'ÊÝÚíá ÒÑ ÇáäÇÞÕ æ ÇáíÓÇæí
If Data1.Recordset.RecordCount <= 0 Then
Command4.Enabled = False
Command3.Enabled = False
Else
Command4.Enabled = True
Command3.Enabled = True
End If

End Function
