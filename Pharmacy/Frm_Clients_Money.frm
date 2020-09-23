VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_Clients_Money 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÍÓÇÈÇÊ ÇáÚãáÇÁ"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Clients_Money.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "ÇáÍÇáÉ ÈÔßá ÚÇã"
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   9720
      Width           =   3855
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   50
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   200
         Width           =   3735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "ãÌãæÚ ÍÞá ÇáÏÇÆä"
      Height          =   615
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   9720
      Width           =   1815
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   50
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ãÌãæÚ ÍÞá ÇáãÏíä"
      Height          =   615
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   9720
      Width           =   2055
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   50
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   200
         Width           =   1935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "Frm_Clients_Money.frx":29C12
      Top             =   4560
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clients"
      RightToLeft     =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÈÍË æ ÊÕÝíÉ"
         Height          =   855
         Left            =   6600
         Picture         =   "Frm_Clients_Money.frx":29E46
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÚÑÖ Çáßá"
         Height          =   855
         Left            =   5520
         Picture         =   "Frm_Clients_Money.frx":2A338
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÐÝ ãÈáÛ"
         Height          =   855
         Left            =   8040
         Picture         =   "Frm_Clients_Money.frx":2A77A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÚÏíá ãÈáÛ"
         Height          =   855
         Left            =   9120
         Picture         =   "Frm_Clients_Money.frx":2ABBC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÖÇÝÉ ãÈáÛ"
         Height          =   855
         Left            =   10200
         Picture         =   "Frm_Clients_Money.frx":2AFFE
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
         Picture         =   "Frm_Clients_Money.frx":2B440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   7935
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         DataField       =   "why"
         DataSource      =   "Data2"
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Text            =   "Text7"
         Top             =   2.45745e5
         Width           =   495
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "clients_money"
         RightToLeft     =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         DataField       =   "date"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   2.45745e5
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         DataField       =   "pdaen"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2.45745e5
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "pmdeon"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   2.45745e5
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "clientcode"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   2.45745e5
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "clientcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   2.45745e5
         Width           =   375
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "Frm_Clients_Money.frx":2BA3C
         Height          =   8415
         Left            =   0
         OleObjectBlob   =   "Frm_Clients_Money.frx":2BA50
         TabIndex        =   10
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   8040
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "clientname"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   65454
         Width           =   615
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Frm_Clients_Money.frx":2C783
         Height          =   8415
         Left            =   0
         OleObjectBlob   =   "Frm_Clients_Money.frx":2C797
         TabIndex        =   9
         Top             =   120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Frm_Clients_Money"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frm_Add_Money.Show
StayOnTop Frm_Add_Money
Frm_Add_Money.Text1.Text = Text1.Text
Frm_Add_Money.Text4.Text = Date
storecommand3 = "addnew"

End Sub

Private Sub Command2_Click()
If Text3.Text <> "" Then
Frm_Add_Money.Show
StayOnTop Frm_Add_Money
storecommand3 = "edit"
Frm_Add_Money.Text1.Text = Text1.Text
Frm_Add_Money.Text2.Text = Text4.Text
Frm_Add_Money.Text3.Text = Text5.Text
Frm_Add_Money.Text4.Text = Text6.Text
Frm_Add_Money.Text5.Text = Text7.Text
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáãÈáÛ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
End If

End Sub

Private Sub Command3_Click()
If Text3.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáãÈáÛ", 16, "äÙÇã ÇáÕíÏáíÉ 2007"
Exit Sub
End If

Dim mok
mok = MsgBox(" åá ÃäÊ ãÊÃßÏ Ãäß ÊÑíÏ ÍÐÝ ÇáãÈáÛ ãä ÏÝÊÑ ÍÓÇÈÇÊ " & Text1.Text, 64 + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ")
If mok = vbYes Then
On Error Resume Next
Data2.Recordset.Delete
Data2.Refresh
DBGrid2.ReBind
DBGrid2.Refresh
refreshall
Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Frm_Clients_Find.Show
StayOnTop Frm_Clients_Find

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from clients"
Data1.Refresh

'ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáËÇäíÉ
Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data2.RecordSource = "select * from clients_money where clientcode=" & Text2.Text & ""
Data2.Refresh
DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & Text1.Text
DBGrid2.Refresh
DBGrid2.ReBind

End Sub

Private Sub DBGrid1_Click()
Data2.RecordSource = "select * from clients_money where clientcode=" & Text2.Text & ""
Data2.Refresh
DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & Text1.Text
DBGrid2.Refresh
DBGrid2.ReBind
refreshall
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hwnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from clients"
Data1.Refresh

'ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáËÇäíÉ
Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data2.RecordSource = "select * from clients_money where clientcode=" & Text2.Text & ""
Data2.Refresh
DBGrid2.Caption = " ÏÝÊÑ ÍÓÇÈ " & Text1.Text
DBGrid2.Refresh
DBGrid2.ReBind

refreshall
End Sub

Public Function refreshall()
Dim i As Integer
On Error Resume Next
Data2.Recordset.MoveFirst
Text8.Text = CDbl(0)
For i = 1 To Data2.Recordset.RecordCount
On Error Resume Next
Text8.Text = CDbl(Text8.Text) + CDbl(Text4.Text)
Data2.Recordset.MoveNext
Next

Dim j As Integer
On Error Resume Next
Data2.Recordset.MoveFirst
Text9.Text = CDbl(0)
For j = 1 To Data2.Recordset.RecordCount
On Error Resume Next
Text9.Text = CDbl(Text9.Text) + CDbl(Text5.Text)
Data2.Recordset.MoveNext
Next

If CDbl(Text9.Text) = CDbl(Text8.Text) Then
Text10.Text = "Êã ÊÕÝíÉ ÇáÍÓÇÈ"
Text10.BackColor = &HFFFFFF
End If
If CDbl(Text9.Text) > CDbl(Text8.Text) Then
Text10.Text = " ÏÇÆä ÚáíäÇ ÈãÈáÛ " & (CDbl(Text9.Text) - CDbl(Text8.Text))
Text10.BackColor = &HC0C0FF
End If
If CDbl(Text9.Text) < CDbl(Text8.Text) Then
Text10.Text = " ãÏíä áäÇ ÈãÈáÛ " & (CDbl(Text8.Text) - CDbl(Text9.Text))
Text10.BackColor = &HC0FFC0
End If

If CDbl(Text8.Text) = CDbl(0) And CDbl(Text9.Text) = CDbl(0) Then
Text10.Text = " ÏÝÊÑ ÇáÍÓÇÈ ÝÇÑÛ "
Text10.BackColor = &HFFFFFF
End If
End Function
