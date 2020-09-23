VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_EndClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅäåÇÁ ÇáãÈíÚ"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_EndClient.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   1200
      OleObjectBlob   =   "Frm_EndClient.frx":29C12
      Top             =   960
   End
   Begin VB.TextBox Text_sum 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   120
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "Frm_EndClient.frx":29E46
      Top             =   2400
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   3720
      OleObjectBlob   =   "Frm_EndClient.frx":2A07A
      TabIndex        =   3
      Top             =   0
      Width           =   1350
   End
End
Attribute VB_Name = "Frm_EndClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text_sum.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÚãíá", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

'ÇÖÇÝÉ ÇáãÈáÛ
With Frm_Clients_Money
.Data2.Recordset.AddNew
.Text3.Text = selcode
.Text4.Text = CDbl(summ1)
.Text5.Text = CDbl(0)
.Text6.Text = Format(Date, "Short Date")
.Text7.Text = "ÔÑÇÁ ÃÏæíÉ ÈÇáÃÌá"
On Error Resume Next
.Data2.Recordset.MoveNext
.Data2.Recordset.MovePrevious
.DBGrid2.Refresh
.DBGrid2.ReBind
End With
Frm_Clients_Money.refreshall
Unload Frm_Clients_Money
'ÊÝÑíÛ ÇáÔÈßÉ ËÇäí Ôí
With frm_SalePoint
Dim I As Integer
On Error Resume Next
.Data1.Recordset.MoveFirst
For I = 1 To .Data1.Recordset.RecordCount
.Data1.Recordset.Delete
.Data1.Recordset.MoveFirst
Next
.DBGrid1.ReBind
.DBGrid1.Refresh
.Text1.Text = ""
.Text2.Text = ""
.Text3.Text = "0"
.Text4.Text = ""
.refreshall
End With

Unload Me



End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Text_sum_DblClick()
frm_list3.Show
StayOnTop frm_list3

End Sub

Private Sub Text_sum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_list3.Show
StayOnTop frm_list3
End If
End Sub
