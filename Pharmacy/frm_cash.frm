VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_cash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáãÈáÛ ÇáãÞÈæÖ ÚäÏ ÇáÅÚÇÏÉ"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   600
      Width           =   9855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3000
      OleObjectBlob   =   "frm_cash.frx":0000
      Top             =   2520
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   495
      Left            =   3240
      OleObjectBlob   =   "frm_cash.frx":0234
      TabIndex        =   3
      Top             =   120
      Width           =   6750
   End
End
Attribute VB_Name = "frm_cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Ýí ÇáÈÏÇíÉ ÅÖÇÝÉ ÇáãÈáÛ  ááÚãíá ÊÍÊ ÇÓã ÇÑÊÌÇÚ áÍÞáí ÇáãÏíä æ ÇáÏÇÆä ãÚÇð
'ÇÖÇÝÉ ÇáãÈáÛ
Dim selcode As Long
'ÈÍË Úä ÇáãæÑÏ Ýí ÌÏæá ÇáÚãáÇÁ
With Frm_Clients
  .Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
  .Data1.RecordSource = "select * from clients where clientname='" & Frm_Disactive.Text7.Text & "'"
  .Data1.Refresh
 
  If .Text1.Text <> "" Then
  selcode = .Text2.Text
  End If
End With

With Frm_Clients_Money
.Data2.Recordset.AddNew
.Text3.Text = selcode
.Text4.Text = CDbl(0)
.Text5.Text = CDbl(Text_sum)
.Text6.Text = Format(Date, "Short Date")
.Text7.Text = "ÞíãÉ ÃÏæíÉ ÝÇÓÏÉ ãÑÊÌÚÉ"
On Error Resume Next
.Data2.Recordset.MoveNext
.Data2.Recordset.MovePrevious
.DBGrid2.Refresh
.DBGrid2.ReBind
End With

  ' ÈÏÁ ÇáÚãáíøÉ
  With frm_store
    .Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
    .Data1.RecordSource = "select * from pharstore where comname='" & Frm_Disactive.Text1.Text & "' and doname='" & Frm_Disactive.Text2.Text & "' and docode='" & Frm_Disactive.Text3.Text & "'"
    .Data1.Refresh
    
   If .Text1.Text <> "" Then
      .Data1.Recordset.Delete
    On Error Resume Next
      .Data1.Recordset.MoveNext
     .Data1.Recordset.MovePrevious
   End If
  End With
  Unload Me
  MsgBox "ÊãÊ ÇáÚãáíÉ ÈäÌÇÍ", 64, "äÙÇã ÇáÕíÏáíøÉ 2007"
  frm_main.Initial

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

