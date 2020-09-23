VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_date2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÑÌÇÁ ÇÏÎÇá ÊÇÑíÎíøä"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_date2.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frm_date2.frx":29C12
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ãÜÜæÇÝÜÜÞ"
         Default         =   -1  'True
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frm_date2.frx":29C86
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "frm_date2.frx":29D00
      Top             =   480
   End
End
Attribute VB_Name = "frm_date2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsDate(Text1.Text) = False Or IsDate(Text2.Text) = False Then
MsgBox "ÇáÑÌÇÁ ÇÏÎÇá ÇáÊæÇÑíÎ ÈÔßá ÕÍíÍ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text1.Text = Format(Date, "Short Date")
Text2.Text = CDate(Text1.Text) + CDate(1)
Exit Sub
End If

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÇÏÎÇá ÇáÊæÇÑíÎ ÈÔßá ÕÍíÍ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Text1.Text = Format(Date, "Short Date")
Text2.Text = CDate(Text1.Text) + CDate(1)
Exit Sub
End If

If CDate(Text2.Text) < CDate(Text1.Text) Then
MsgBox "íÌÈ Ãä íßæä ÊÇÑíÎ ÇáÅäÊåÇÁ ÈÚÏ ÊÇÑíÎ ÇáÈÏÁ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
'ÈÏÇíÉ ÇáÊØÈíÞ
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
           .Commands(2).CommandType = adCmdText
          .Commands(2).CommandText = "select * from daily where date between #" & CDate(Format(Text1.Text, "Short Date")) & "# and #" & CDate(Format(Text2.Text, "Short Date")) & "#"
           .Commands(2).Execute
            
         If .rsCommand2.State = 1 Then
         
           .rsCommand2.Close
         
         End If
         
End With
DataReport2.Sections(2).Controls("label1").Caption = "ãÈíÚÇÊ ÇáÕíÏáíøÉ Èíä " & Format(Text1.Text, "Short Date") & " æ " & Format(Text2.Text, "Short Date")
DataReport2.Sections(2).Controls("Label2").Caption = " ÊÇÑíÎ ÇÕÏÇÑ ÇáÊÞÑíÑ " & Format(Date, "Short Date") & " - " & " ÇáÓÇÚÉ : " & Time
DataReport2.Caption = "ÊÞÑíÑ Úä ãÈíÚÇÊ ÇáÕíÏáíøÉ Èíä " & Format(Text1.Text, "Short Date") & " æ " & Format(Text2.Text, "Short Date")
'ÍÓÇÈ ãÌãæÚ ÇÌãÇáí ÇáãÈíÚÇÊ
Dim I As Long
Dim S As Double
S = CDbl(0)
With frm_daily
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from daily where date between #" & CDate(Format(Text1.Text, "Short Date")) & "# and #" & CDate(Format(Text2.Text, "Short Date")) & "#"
.Data1.Refresh
On Error Resume Next
.Data1.Recordset.MoveFirst
For I = 1 To .Data1.Recordset.RecordCount
S = CDbl(S) + CDbl(.Text1.Text)
.Data1.Recordset.MoveNext
Next
End With
Unload frm_daily
DataReport2.Sections(5).Controls("Label3").Caption = CDbl(S)
DataReport2.Sections(5).Controls("Label6").Caption = " äÙÇã ÇáÕíÏáíøÉ 2007 -  " & mypharname

S = Empty
I = Empty

DataReport2.Show
Unload Me






End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

End Sub

