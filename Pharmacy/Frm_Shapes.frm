VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Shapes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ôßá ÇáÚÈæÉ"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "shapeof"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ Ôßá"
         Height          =   855
         Left            =   1560
         Picture         =   "Frm_Shapes.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá Ôßá"
         Height          =   855
         Left            =   2520
         Picture         =   "Frm_Shapes.frx":0442
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ôßá ÌÏíÏ"
         Height          =   855
         Left            =   3480
         Picture         =   "Frm_Shapes.frx":0884
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Shapes.frx":0CC6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "shapes"
      RightToLeft     =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Frm_Shapes.frx":12C2
      Top             =   2760
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_Shapes.frx":14F6
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "Frm_Shapes.frx":150A
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
End
Attribute VB_Name = "Frm_Shapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frm_Addshape.Show
Frm_Addshape.Caption = "ÅÖÇÝÉ Ôßá ÚÈæÉ ÌÏíÏ"
StayOnTop Frm_Addshape
shapecommand1 = "addnew"

End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
Frm_Addshape.Show
StayOnTop Frm_Addshape
shapecommand1 = "edit"
Frm_Addshape.Text1.Text = Text1.Text
Frm_Addshape.Caption = "ÊÚÏíá Ôßá ÚÈæÉ"
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ Ôßá ÇáÚÈæÉ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ Ôßá ÇáÚÈæÉ ÇáãØáæÈÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ Ôßá ÇáÚÈæÉ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Else
Exit Sub
End If
frm_main.Refreshcommand

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hwnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from shapes"
Data1.Refresh

End Sub
