VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Disactive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÃÏæíÉ ÐÇÊ ÇáÝÚÇáíøÉ ÇáãäÊåíÉ"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      DataField       =   "storagewh"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text9"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      DataField       =   "pharprice"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text8"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "mordname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "shape"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "docount"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "peiceprice"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "docode"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "doname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comname"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÚÇÏÉ ááãæÑøÏ"
         Height          =   855
         Left            =   6960
         Picture         =   "Frm_Disactive.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊãÏíÏ ÇáÝÚøÇáíøÉ"
         Height          =   855
         Left            =   9120
         Picture         =   "Frm_Disactive.frx":06D8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Disactive.frx":0E5E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÊáÇÝ ÇáÏæÇÁ"
         Height          =   855
         Left            =   8040
         Picture         =   "Frm_Disactive.frx":145A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "disactive"
      RightToLeft     =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4800
      OleObjectBlob   =   "Frm_Disactive.frx":1B66
      Top             =   2880
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_Disactive.frx":1D9A
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "Frm_Disactive.frx":1DAE
      TabIndex        =   4
      Top             =   1080
      Width           =   10215
   End
End
Attribute VB_Name = "Frm_Disactive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim mok
mok = MsgBox("ÅÐÇ ÇÎÊÑÊ äÚã ÓæÝ íÊã ÍÐÝ ÇáÏæÇÁ ãä ÇáãÓÊæÏÚ ÈÔßá äåÇÆí åá ÊÑíÏ ÇáãÊÇÈÚÉ ¿", 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
  'ÈÏÁ ÇáÚãáíøÉ
  With frm_store
    .Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
    .Data1.RecordSource = "select * from pharstore where comname='" & Text1.Text & "' and doname='" & Text2.Text & "' and docode='" & Text3.Text & "'"
    .Data1.Refresh
    
   If .Text1.Text <> "" Then
      .Data1.Recordset.Delete
    On Error Resume Next
      .Data1.Recordset.MoveNext
     .Data1.Recordset.MovePrevious
   End If
  End With
  MsgBox "ÊãÊ ÇáÚãáíÉ ÈäÌÇÍ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
  frm_main.Initial
Else
  Exit Sub
End If
End Sub

Private Sub Command2_Click()
Frm_Increment.Show
StayOnTop Frm_Increment
Frm_Increment.Text1.Text = Text1.Text
Frm_Increment.Text2.Text = Text2.Text
Frm_Increment.Text3.Text = Text3.Text
Frm_Increment.Text5.SetFocus
End Sub

Private Sub Command3_Click()
frm_cash.Show
StayOnTop frm_cash

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from disactive"
Data1.Refresh


End Sub
