VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_reports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÊÞÇÑíÑ"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_reports.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÊÞÇÑíÑ ÇáãÇáíøÉ"
         Height          =   855
         Left            =   5040
         Picture         =   "frm_reports.frx":29C12
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÞÇÑíÑ ÇáÚãáÇÁ"
         Height          =   855
         Left            =   6480
         Picture         =   "frm_reports.frx":2A326
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÊÞÇÑíÑ ÇáãÓÊæÏÚ"
         Height          =   855
         Left            =   7920
         Picture         =   "frm_reports.frx":2AAB2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "frm_reports.frx":2B288
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4560
      OleObjectBlob   =   "frm_reports.frx":2B884
      Top             =   3360
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÍÓÇÈ Úãíá"
         Height          =   855
         Left            =   120
         Picture         =   "frm_reports.frx":2BAB8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÃÏæíÉ ÇáÝÇÓÏÉ"
         Height          =   855
         Left            =   120
         Picture         =   "frm_reports.frx":2C10C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÃÏæíÉ ÇáäÇÝÐÉ"
         Height          =   855
         Left            =   1440
         Picture         =   "frm_reports.frx":2C925
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇáÃÏæíÉ ÇáãÊæÝÑÉ"
         Height          =   855
         Left            =   2760
         Picture         =   "frm_reports.frx":2D13E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ãÈíÚÇÊ Èíä ÊÇÑíÎíä"
         Height          =   855
         Left            =   120
         Picture         =   "frm_reports.frx":2D88F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ãÈíÚÇÊ ÊÇÑíÎ ãÍÏÏ"
         Height          =   855
         Left            =   1560
         Picture         =   "frm_reports.frx":2E01B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇÌãÇáí ÇáãÈíÚÇÊ"
         Height          =   855
         Left            =   3000
         Picture         =   "frm_reports.frx":2E7AF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
           .Commands(1).CommandType = adCmdText
          .Commands(1).CommandText = "select * from pharstore where docount>0"
           .Commands(1).Execute
            
         If .rsCommand1.State = 1 Then
         
           .rsCommand1.Close
         
         End If
         
End With
DataReport1.Sections(2).Controls("label1").Caption = "ÇáÃÏæíÉ ÇáãÊæÝÑÉ Ýí ÇáÕíÏáíøÉ"
DataReport1.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport1.Sections(5).Controls("Label3").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname

DataReport1.Caption = "ÊÞÑíÑ Úä ÇáÃÏæíÉ ÇáãÊæÝÑÉ Ýí ÇáÕíÏáíøÉ"
DataReport1.Show
End Sub

Private Sub Command2_Click()
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
           .Commands(1).CommandType = adCmdText
          .Commands(1).CommandText = "select * from pharstore where docount=0"
           .Commands(1).Execute
            
         If .rsCommand1.State = 1 Then
         
           .rsCommand1.Close
         
         End If
         
End With
DataReport1.Sections(2).Controls("label1").Caption = "ÇáÃÏæíÉ ÇáäÇÝÐÉ ãä ÇáÕíÏáíøÉ"
DataReport1.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport1.Sections(5).Controls("Label3").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname

DataReport1.Caption = "ÊÞÑíÑ Úä ÇáÃÏæíÉ ÇáäÇÝÐÉ ãä ÇáÕíÏáíøÉ"
DataReport1.Show

End Sub

Private Sub Command3_Click()
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
frm_store.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
frm_store.Data1.RecordSource = "select * from pharstore"
frm_store.Data1.Refresh
 On Error Resume Next
 frm_store.Data1.Recordset.MoveFirst

If frm_store.Text5.Text = "" Then
Exit Sub
Else
 Dim i As Long
  'ÊÍÖíÑ ÇáÈÑæÌÑÓ ÈÇÑ
 FrmProgress.bar.Max = frm_store.Data1.Recordset.RecordCount
 FrmProgress.bar.Min = 0
 FrmProgress.bar.Value = 0
 FrmProgress.Show
 StayOnTop FrmProgress
 For i = 1 To frm_store.Data1.Recordset.RecordCount
  If (CDate(frm_store.Text11.Text) - CDate(Date)) < sharp Then
   With Frm_Disactive
    .Data1.Recordset.AddNew
    .Text1.Text = frm_store.Text1.Text
    .Text2.Text = frm_store.Text2.Text
    .Text3.Text = frm_store.Text3.Text
    .Text4.Text = CDbl(frm_store.Text4.Text)
    .Text5.Text = CLng(frm_store.Text5.Text)
    .Text6.Text = frm_store.Text6.Text
    .Text7.Text = frm_store.Text7.Text
    .Text8.Text = CDbl(frm_store.Text8.Text)
    .Text9.Text = frm_store.Text10.Text
    .Data1.Recordset.MoveNext
    .Data1.Recordset.MovePrevious

   End With
  End If
  FrmProgress.bar.Value = i
  frm_store.Data1.Recordset.MoveNext
 Next
 Unload FrmProgress
End If
Unload Frm_Disactive
'ÚÑÖ ÇáÞÑíÑ ÇáÂä
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
           .Commands(3).CommandType = adCmdText
          .Commands(3).CommandText = "select * from disactive"
           .Commands(3).Execute
            
         If .rsCommand3.State = 1 Then
         
           .rsCommand3.Close
         
         End If
         
End With

DataReport4.Sections(2).Controls("label1").Caption = "ÇáÃÏæíÉ ÐÇÊ ÇáÕáÇÍíøÉ ÇáãäÊåíÉ"
DataReport4.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport4.Sections(5).Controls("Label3").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname
DataReport4.Caption = "ÊÞÑíÑ Úä ÇáÃÏæíÉ ÇáÝÇÓÏÉ Ýí ÇáÕíÏáíøÉ"
DataReport4.Show



End Sub

Private Sub Command4_Click()
Frm_client_cash.Show
StayOnTop Frm_client_cash
End Sub

Private Sub Command5_Click()
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
           .Commands(2).CommandType = adCmdText
          .Commands(2).CommandText = "select * from daily order by date"
           .Commands(2).Execute
            
         If .rsCommand2.State = 1 Then
         
           .rsCommand2.Close
         
         End If
         
End With
DataReport2.Sections(2).Controls("label1").Caption = "ÇÌãÇáí ãÈíÚÇÊ ÇáÕíÏáíøÉ"
DataReport2.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport2.Caption = "ÊÞÑíÑ Úä ÇÌãÇáí ãÈíÚÇÊ ÇáÕíÏáíøÉ"
'ÍÓÇÈ ãÌãæÚ ÇÌãÇáí ÇáãÈíÚÇÊ
Dim i As Long
Dim S As Double
S = CDbl(0)
With frm_daily
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from daily"
.Data1.Refresh
On Error Resume Next
.Data1.Recordset.MoveFirst
For i = 1 To .Data1.Recordset.RecordCount
S = CDbl(S) + CDbl(.Text1.Text)
.Data1.Recordset.MoveNext
Next
End With
Unload frm_daily
DataReport2.Sections(5).Controls("Label3").Caption = CDbl(S)
DataReport2.Sections(5).Controls("Label6").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname

S = Empty
i = Empty

DataReport2.Show

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Command7_Click()
frm_date1.Text1.Text = Format(Date, "Short Date")
frm_date1.Show
StayOnTop frm_date1

End Sub

Private Sub Command8_Click()
frm_date2.Text1.Text = Format(Date, "Short Date")
frm_date2.Text2.Text = CDate(frm_date2.Text1.Text) + CDate(1)
frm_date2.Show
StayOnTop frm_date2
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

Public Function refreshframes()
If Option1.Value = True Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If

If Option2.Value = True Then
Frame3.Visible = True
Else
Frame3.Visible = False
End If

If Option3.Value = True Then
Frame4.Visible = True
Else
Frame4.Visible = False
End If


End Function

Private Sub Option1_Click()
Me.refreshframes
End Sub

Private Sub Option2_Click()
Me.refreshframes

End Sub

Private Sub Option3_Click()
Me.refreshframes

End Sub
