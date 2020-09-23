VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_store 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÓÊæÏÚ ÇáÕíÏáíøÉ"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_store.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      DataField       =   "doactivity"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Text            =   "Text11"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "storagewh"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Text            =   "Text10"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      DataField       =   "pharprice"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "mordname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "shape"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "docount"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "peiceprice"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2.45745e5
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "docode"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "doname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_store.frx":29C12
      Height          =   8055
      Left            =   120
      OleObjectBlob   =   "frm_store.frx":29C26
      TabIndex        =   9
      Top             =   1080
      Width           =   11775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pharstore"
      RightToLeft     =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5760
      OleObjectBlob   =   "frm_store.frx":2AB08
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ØÈÇÚÉ ÇááÕÇÞÇÊ"
         Height          =   855
         Left            =   1560
         Picture         =   "frm_store.frx":2AD3C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÃÏæíÉ ÇáÝÇÓÏÉ"
         Height          =   855
         Left            =   2880
         Picture         =   "frm_store.frx":2B17E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ãÔÇÈå"
         Height          =   855
         Left            =   9600
         Picture         =   "frm_store.frx":2B7FD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÚÑÖ Çáßá"
         Height          =   855
         Left            =   5280
         Picture         =   "frm_store.frx":2BC3F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "frm_store.frx":2C081
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÃÏæíÉ ÇáäÇÝÐÉ"
         Height          =   855
         Left            =   4080
         Picture         =   "frm_store.frx":2C67D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÈÍË æ ÊÕÝíÉ"
         Height          =   855
         Left            =   6360
         Picture         =   "frm_store.frx":2CABF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ÏæÇÁ"
         Height          =   855
         Left            =   10680
         Picture         =   "frm_store.frx":2CFB1
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá ÏæÇÁ"
         Height          =   855
         Left            =   8520
         Picture         =   "frm_store.frx":2D3F3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ ÏæÇÁ"
         Height          =   855
         Left            =   7440
         Picture         =   "frm_store.frx":2D835
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frm_AddDoa.Show
StayOnTop Frm_AddDoa
storecommand1 = "addnew"
End Sub

Private Sub Command10_Click()
If Text1.Text <> "" Then
With DataReport3.Sections("section1")

.Controls("label1").Caption = "* " & Text3.Text & " *"
.Controls("name1").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label2").Caption = "* " & Text3.Text & " *"
.Controls("name2").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label3").Caption = "* " & Text3.Text & " *"
.Controls("name3").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label4").Caption = "* " & Text3.Text & " *"
.Controls("name4").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label5").Caption = "* " & Text3.Text & " *"
.Controls("name5").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label6").Caption = "* " & Text3.Text & " *"
.Controls("name6").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label7").Caption = "* " & Text3.Text & " *"
.Controls("name7").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label8").Caption = "* " & Text3.Text & " *"
.Controls("name8").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label9").Caption = "* " & Text3.Text & " *"
.Controls("name9").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label10").Caption = "* " & Text3.Text & " *"
.Controls("name10").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label11").Caption = "* " & Text3.Text & " *"
.Controls("name11").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label12").Caption = "* " & Text3.Text & " *"
.Controls("name12").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label13").Caption = "* " & Text3.Text & " *"
.Controls("name13").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label14").Caption = "* " & Text3.Text & " *"
.Controls("name14").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label15").Caption = "* " & Text3.Text & " *"
.Controls("name15").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label16").Caption = "* " & Text3.Text & " *"
.Controls("name16").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label17").Caption = "* " & Text3.Text & " *"
.Controls("name17").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label18").Caption = "* " & Text3.Text & " *"
.Controls("name18").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label19").Caption = "* " & Text3.Text & " *"
.Controls("name19").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label20").Caption = "* " & Text3.Text & " *"
.Controls("name20").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label21").Caption = "* " & Text3.Text & " *"
.Controls("name21").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label22").Caption = "* " & Text3.Text & " *"
.Controls("name22").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label23").Caption = "* " & Text3.Text & " *"
.Controls("name23").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label24").Caption = "* " & Text3.Text & " *"
.Controls("name24").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label25").Caption = "* " & Text3.Text & " *"
.Controls("name25").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label26").Caption = "* " & Text3.Text & " *"
.Controls("name26").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label27").Caption = "* " & Text3.Text & " *"
.Controls("name27").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label28").Caption = "* " & Text3.Text & " *"
.Controls("name28").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label29").Caption = "* " & Text3.Text & " *"
.Controls("name29").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label30").Caption = "* " & Text3.Text & " *"
.Controls("name30").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label31").Caption = "* " & Text3.Text & " *"
.Controls("name31").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label32").Caption = "* " & Text3.Text & " *"
.Controls("name32").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label33").Caption = "* " & Text3.Text & " *"
.Controls("name33").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label34").Caption = "* " & Text3.Text & " *"
.Controls("name34").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label35").Caption = "* " & Text3.Text & " *"
.Controls("name35").Caption = mypharname & vbNewLine & Text2.Text


.Controls("label36").Caption = "* " & Text3.Text & " *"
.Controls("name36").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label39").Caption = "* " & Text3.Text & " *"
.Controls("name39").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label40").Caption = "* " & Text3.Text & " *"
.Controls("name40").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label41").Caption = "* " & Text3.Text & " *"
.Controls("name41").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label42").Caption = "* " & Text3.Text & " *"
.Controls("name42").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label43").Caption = "* " & Text3.Text & " *"
.Controls("name43").Caption = mypharname & vbNewLine & Text2.Text

.Controls("label44").Caption = "* " & Text3.Text & " *"
.Controls("name44").Caption = mypharname & vbNewLine & Text2.Text

End With
DataReport3.Caption = "ãÚÇíäÉ ÞÈá ØÈÇÚÉ ÇááÕÇÞÇÊ"

DataReport3.Show
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"

End If
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
Frm_AddDoa.Show
StayOnTop Frm_AddDoa
storecommand1 = "edit"
Frm_AddDoa.Text1.Text = Text1.Text
Frm_AddDoa.Text2.Text = Text2.Text
Frm_AddDoa.Text3.Text = Text3.Text
Frm_AddDoa.Text4.Text = Text4.Text
Frm_AddDoa.Text5.Text = Text5.Text
Frm_AddDoa.Text7.Text = Text6.Text
Frm_AddDoa.Text8.Text = Text7.Text
Frm_AddDoa.Text6.Text = Text8.Text
Frm_AddDoa.Text9.Text = Text10.Text
Frm_AddDoa.Text10.Text = Text11.Text

Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÏæÇÁ " & Text2.Text & " ãõäÊÌ ÔÑßÉ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()
frm_store_find.Show
StayOnTop frm_store_find
End Sub

Private Sub Command5_Click()
'ÍÐÝ ÌãíÚ ÓÌáÇÊ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÊÈÚ ÇáÃæíÉ ÇáäÇÝÐÉ
With Frm_Naf
Dim X As Integer
On Error Resume Next
.Data1.Recordset.MoveFirst
For X = 1 To .Data1.Recordset.RecordCount
.Data1.Recordset.Delete
.Data1.Recordset.MoveFirst
Next
End With
'ÊÚÈÆÉ ÇáÞÇÚÏÉ ÈÇáÃÏæíÉ ÇáäÇÝÐÉ ÝÚáÇð
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharstore"
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
 On Error Resume Next
 Data1.Recordset.MoveFirst

If Text5.Text = "" Then
MsgBox "áÇ íæÌÏ Ãí ÃÏæíÉ Ýí ÇáãÓÊæÏÚ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
Else
 Dim I As Long
 'ÊÍÖíÑ ÇáÈÑæÌÑÓ ÈÇÑ
 FrmProgress.bar.Max = Data1.Recordset.RecordCount
 FrmProgress.bar.Min = 0
 FrmProgress.bar.Value = 0
 FrmProgress.Show
 StayOnTop FrmProgress
 For I = 1 To Data1.Recordset.RecordCount
  If CInt(Text5.Text) = CInt(0) Then
   With Frm_Naf
    .Data1.Recordset.AddNew
    .Text1.Text = Text1.Text
    .Text2.Text = Text2.Text
    .Text3.Text = Text3.Text
    .Text4.Text = CDbl(Text4.Text)
    .Text5.Text = CLng(Text5.Text)
    .Data1.Recordset.MoveNext
    .Data1.Recordset.MovePrevious

   End With
  End If
  FrmProgress.bar.Value = I
  Data1.Recordset.MoveNext
 Next
Unload FrmProgress
If Frm_Naf.Data1.Recordset.RecordCount > Int(0) Then
 Frm_Naf.Show
 StayOnTop Frm_Naf
 Else
  MsgBox "áÇ íæÌÏ Ãí ÃÏæíÉ äÇÝÐÉ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
End If

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharstore"
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Me.simi
End Sub

Private Sub Command8_Click()
Frm_AddDoa.Show
StayOnTop Frm_AddDoa
storecommand1 = "addnew"
Frm_AddDoa.Text1.Text = Text1.Text
Frm_AddDoa.Text7.Text = Text6.Text
Frm_AddDoa.Text8.Text = Text7.Text
Frm_AddDoa.Text9.Text = Text10.Text

End Sub

Private Sub Command9_Click()
'ÍÐÝ ÌãíÚ ÓÌáÇÊ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÊÈÚ ÇáÃæíÉ ÇáÝÇÓÏÉ
With Frm_Disactive
Dim X As Integer
On Error Resume Next
.Data1.Recordset.MoveFirst
For X = 1 To .Data1.Recordset.RecordCount
.Data1.Recordset.Delete
.Data1.Recordset.MoveFirst
Next
End With

'ÊÚÈÆÉ ÇáÞÇÚÏÉ ÈÇáÃÏæíÉ ÇáÝÇÓÏÉ ÝÚáÇð
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharstore"
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
 On Error Resume Next
 Data1.Recordset.MoveFirst

If Text5.Text = "" Then
MsgBox "áÇ íæÌÏ Ãí ÃÏæíÉ Ýí ÇáãÓÊæÏÚ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
Else
 Dim I As Long
  'ÊÍÖíÑ ÇáÈÑæÌÑÓ ÈÇÑ
 FrmProgress.bar.Max = Data1.Recordset.RecordCount
 FrmProgress.bar.Min = 0
 FrmProgress.bar.Value = 0
 FrmProgress.Show
 StayOnTop FrmProgress
 For I = 1 To Data1.Recordset.RecordCount
  If (CDate(Text11.Text) - CDate(Date)) < sharp Then
   With Frm_Disactive
    .Data1.Recordset.AddNew
    .Text1.Text = Text1.Text
    .Text2.Text = Text2.Text
    .Text3.Text = Text3.Text
    .Text4.Text = CDbl(Text4.Text)
    .Text5.Text = CLng(Text5.Text)
    .Text6.Text = Text6.Text
    .Text7.Text = Text7.Text
    .Text8.Text = CDbl(Text8.Text)
    .Text9.Text = Text10.Text
    .Data1.Recordset.MoveNext
    .Data1.Recordset.MovePrevious

   End With
  End If
  FrmProgress.bar.Value = I
  Data1.Recordset.MoveNext
 Next
 Unload FrmProgress
If Frm_Disactive.Data1.Recordset.RecordCount > Int(0) Then
 Frm_Disactive.Show
 StayOnTop Frm_Disactive
 Else
  MsgBox "áÇ íæÌÏ Ãí ÃÏæíÉ ÝÇÓÏÉ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
End If

End Sub

Private Sub DBGrid1_Click()
Me.simi
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.simi

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharstore"
Data1.Refresh
Me.simi
End Sub

Public Function simi()
If Text1.Text <> "" Then
Command8.Enabled = True
Else
Command8.Enabled = False
End If
End Function
