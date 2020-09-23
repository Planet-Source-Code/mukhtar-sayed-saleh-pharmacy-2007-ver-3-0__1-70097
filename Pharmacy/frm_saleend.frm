VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_saleend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅäåÇÁ ÇáãÈíÚ"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_saleend.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3000
      OleObjectBlob   =   "frm_saleend.frx":29C12
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   5655
   End
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   480
      Width           =   9855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   495
      Left            =   7200
      OleObjectBlob   =   "frm_saleend.frx":29E46
      TabIndex        =   1
      Top             =   0
      Width           =   2790
   End
End
Attribute VB_Name = "frm_saleend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'ÇÖÇÝÉ ÇáÓÌá ááíæãíÉ Çæá Ôí
With frm_daily
.Data1.Recordset.AddNew
.Text1.Text = CDbl(Text_sum.Text)
.Text2.Text = CDate(Format(Date, "Short Date"))
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.DBGrid1.ReBind
.DBGrid1.Refresh
End With
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
