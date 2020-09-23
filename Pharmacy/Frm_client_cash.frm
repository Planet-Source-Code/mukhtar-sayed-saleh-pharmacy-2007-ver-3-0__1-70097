VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_client_cash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÞÑíÑ ÍÓÇÈ Úãíá"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_client_cash.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÜÜæÇÝÜÜÜÞ"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1440
         OleObjectBlob   =   "Frm_client_cash.frx":29C12
         Top             =   600
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "Frm_client_cash.frx":29E46
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frm_client_cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÚãíá ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
'ÝáÊÑÉ Ãæá Ôí ÑÕíÏ ÇáÚãíá
With Frm_Clients_Money
.Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data2.RecordSource = "select * from clients_money where clientcode=" & selcode2 & ""
.Data2.Refresh
.refreshall
'ãáÆ ÇáÊÞÑíÑ
DataReport5.Sections(2).Controls("label1").Caption = " ÍÓÇÈ ÇáÚãíá : " & Text1.Text
DataReport5.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport5.Sections("section1").Controls("Label3").Caption = .Text8.Text
DataReport5.Sections("section1").Controls("Label7").Caption = .Text9.Text
DataReport5.Sections("section1").Controls("Label8").Caption = .Text10.Text
DataReport5.Sections("section5").Controls("Label9").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname
End With
Unload Frm_Clients_Money
Unload Me
DataReport5.Show

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Text1_DblClick()
Frm_List5.Show
StayOnTop Frm_List5

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Frm_List5.Show
StayOnTop Frm_List5
End If
End Sub
