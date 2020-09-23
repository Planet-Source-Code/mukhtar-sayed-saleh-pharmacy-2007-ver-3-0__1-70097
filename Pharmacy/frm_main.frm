VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{14C68610-D6C3-4F35-BBBF-CF69FA56A94E}#1.0#0"; "ClockSHatem.ocx"
Begin VB.Form frm_main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰Ÿ«„ «·’Ìœ·Ì¯… 2007"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frm_main.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   120
      Picture         =   "frm_main.frx":29C12
      RightToLeft     =   -1  'True
      ScaleHeight     =   4395
      ScaleWidth      =   5955
      TabIndex        =   21
      Top             =   6120
      Width           =   5955
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   9360
      Picture         =   "frm_main.frx":2DBDE
      RightToLeft     =   -1  'True
      ScaleHeight     =   1515
      ScaleWidth      =   5835
      TabIndex        =   20
      Top             =   4200
      Width           =   5835
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7440
      Width           =   2535
      Begin ClockS.HatemClocks HatemClocks1 
         Height          =   2310
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   4075
         SecondArmColor  =   16777215
         SecondArmColor  =   16777215
         MinuteArmColor  =   4194304
         HourArmColor    =   12582912
         BackColor       =   -2147483633
         NumberColor     =   16711680
         CaptionColor    =   -2147483630
         Caption         =   ""
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   6360
   End
   Begin VB.Frame Frame2 
      Height          =   375
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   10080
      Width           =   5655
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "frm_main.frx":30A91
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frm_main.frx":30B0B
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_main.frx":30B85
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "√‘ﬂ«· «·⁄»Ê…"
         Height          =   1095
         Left            =   11640
         Picture         =   "frm_main.frx":30BF3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Õ”«»«  «·⁄„·«¡"
         Height          =   1095
         Left            =   6840
         Picture         =   "frm_main.frx":3146F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Õ”«»«  «·„” Œœ„Ì‰"
         Height          =   1095
         Left            =   3240
         Picture         =   "frm_main.frx":31DDA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "‰”Œ «Õ Ì«ÿÌ"
         Height          =   1095
         Left            =   2160
         Picture         =   "frm_main.frx":324EE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "«·„”«⁄œ…"
         Height          =   1095
         Left            =   1080
         Picture         =   "frm_main.frx":32EA9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÕÊ· «·‰Ÿ«„"
         Height          =   1095
         Left            =   0
         Picture         =   "frm_main.frx":33731
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "«· ﬁ«—Ì—"
         Height          =   1095
         Left            =   4440
         Picture         =   "frm_main.frx":33E56
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "«·ÌÊ„Ì¯…"
         Height          =   1095
         Left            =   5640
         Picture         =   "frm_main.frx":34824
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ﬁ«∆„… «·√œÊÌ…"
         Height          =   1095
         Left            =   9240
         Picture         =   "frm_main.frx":35179
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "”Ã· «·⁄„·«¡"
         Height          =   1095
         Left            =   8040
         Picture         =   "frm_main.frx":35A75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "‘—ﬂ«  «·√œÊÌ…"
         Height          =   1095
         Left            =   10440
         Picture         =   "frm_main.frx":36291
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "«·„Ê—¯œÌ‰ ··√œÊÌ…"
         Height          =   1095
         Left            =   12840
         Picture         =   "frm_main.frx":366D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "≈⁄œ«œ«  «·‰Ÿ«„"
         Height          =   1095
         Left            =   14040
         Picture         =   "frm_main.frx":36E63
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "frm_main.frx":3762D
      Top             =   1320
   End
   Begin VB.Menu mnuhelp 
      Caption         =   " ⁄·Ì„« "
      Begin VB.Menu mnu_hlpcontent 
         Caption         =   "„Ê«÷Ì⁄ «· ⁄·Ì„« "
      End
      Begin VB.Menu hhhh 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_about 
         Caption         =   "ÕÊ· «·‰Ÿ«„"
      End
   End
   Begin VB.Menu Mnu_setting 
      Caption         =   "≈⁄œ«œ« "
      Begin VB.Menu userss 
         Caption         =   "Õ”«»«  «·„” Œœ„Ì‰"
      End
      Begin VB.Menu Security 
         Caption         =   "Õ„«Ì… «·‰Ÿ«„"
         Visible         =   0   'False
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu BackUp 
         Caption         =   "«·‰”Œ «·≈Õ Ì«ÿÌ"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu Settingss 
         Caption         =   "≈⁄œ«œ«  «·‰Ÿ«„"
      End
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "√œÊ« "
      Begin VB.Menu Mnu_Morden 
         Caption         =   "«·„Ê—¯œÌ‰ ··√œÊÌ…"
      End
      Begin VB.Menu Mnu_Shapes 
         Caption         =   "√‘ﬂ«· «·⁄»Ê…"
      End
      Begin VB.Menu mnu_companies 
         Caption         =   "‘—ﬂ«  «·√œÊÌ…"
      End
      Begin VB.Menu asas 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_store 
         Caption         =   "„” Êœ⁄ «·’Ìœ·Ì¯…"
      End
      Begin VB.Menu aaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_clients 
         Caption         =   "”Ã· «·⁄„·«¡"
      End
      Begin VB.Menu mnu_clients_money 
         Caption         =   "Õ”«»«  «·⁄„·«¡"
      End
      Begin VB.Menu bbb 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_daily 
         Caption         =   "«·ÌÊ„Ì¯…"
      End
      Begin VB.Menu ccc 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_reports 
         Caption         =   "«· ﬁ«—Ì—"
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "„·›"
      Begin VB.Menu Mne_File_Exit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BackUp_Click()
frm_backup.Show
StayOnTop frm_backup

End Sub

Private Sub Command1_Click()
frmsetting.Show
StayOnTop frmsetting

End Sub

Private Sub Command10_Click()
frm_backup.Show
StayOnTop frm_backup

End Sub

Private Sub Command11_Click()
Frm_Accounts.Show
StayOnTop Frm_Accounts

End Sub

Private Sub Command12_Click()
Frm_Clients_Money.Show
StayOnTop Frm_Clients_Money

End Sub

Private Sub Command13_Click()
Frm_Shapes.Show
StayOnTop Frm_Shapes

End Sub

Private Sub Command2_Click()
Frm_Mord.Show
StayOnTop Frm_Mord
End Sub

Private Sub Command3_Click()
frm_companies.Show
StayOnTop frm_companies
End Sub

Private Sub Command4_Click()
Frm_Clients.Show
StayOnTop Frm_Clients

End Sub

Private Sub Command5_Click()
frm_store.Show
End Sub

Private Sub Command6_Click()
Me.Refreshcommand
frm_daily.Show
StayOnTop frm_daily

End Sub

Private Sub Command7_Click()
frm_reports.Show

End Sub

Private Sub Command8_Click()
frm_about.Show
StayOnTop frm_about
End Sub

Private Sub Command9_Click()
Frm_Help.Show
StayOnTop Frm_Help

End Sub

Private Sub Form_Load()

Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd
Refreshcommand
LoadSetting
readsetting (True)
Initial
'«· ⁄·Ìﬁ« 
If tip = True Then
Frm_tip.Show
StayOnTop Frm_tip
End If

'—»ÿ «·œ« « ≈‰›«Ì—Ê‰„Ì‰ 
With DataEnvironment1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\pharmokhtar.dll;Mode=Read|Write"
End With

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetOrigRes

End Sub

Private Sub Mne_File_Exit_Click()
Dim mok
mok = MsgBox("Â·  —Ìœ »«· √ﬂÌœ «·Œ—ÊÃ „‰ «·‰Ÿ«„ ø", vbYesNo + vbQuestion, "‰Ÿ«„ «·’Ìœ·Ì¯… 2007")
If mok = vbYes Then
Unload Me
End
Else
Exit Sub
End If
End Sub

Private Sub Mnu_about_Click()
frm_about.Show
StayOnTop frm_about

End Sub

Private Sub mnu_clients_Click()
Frm_Clients.Show
StayOnTop Frm_Clients

End Sub

Private Sub mnu_clients_money_Click()
Frm_Clients_Money.Show
StayOnTop Frm_Clients_Money

End Sub

Private Sub mnu_companies_Click()
frm_companies.Show
StayOnTop frm_companies

End Sub

Private Sub mnu_daily_Click()
frm_daily.Show
StayOnTop frm_daily
End Sub

Private Sub mnu_hlpcontent_Click()
Frm_Help.Show
StayOnTop Frm_Help

End Sub

Private Sub Mnu_Morden_Click()
Frm_Mord.Show
StayOnTop Frm_Mord

End Sub

Private Sub mnu_reports_Click()
frm_reports.Show
StayOnTop frm_reports

End Sub

Private Sub Mnu_Shapes_Click()
Frm_Shapes.Show
StayOnTop Frm_Shapes

End Sub

Private Sub mnu_store_Click()
frm_store.Show
StayOnTop frm_store

End Sub

Private Sub Settingss_Click()
frmsetting.Show
StayOnTop frmsetting

End Sub

Private Sub Timer1_Timer()
SkinLabel1.Caption = Time
SkinLabel2.Caption = Date

End Sub

Public Function Refreshcommand()
Dim cond1 As Boolean, cond2 As Boolean, cond3 As Boolean
'«· Õﬁ „‰ ÊÃÊœ ‘—ﬂ«  √œÊÌ…
With frm_companies
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from companies"
.Data1.Refresh
If .Text2.Text <> "" Then
cond1 = True
Else
cond1 = False
End If
End With
'«·‰Õﬁﬁ „‰ ÊÃÊœ √‘ﬂ«· ⁄»Ê…
With Frm_Shapes
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from shapes"
.Data1.Refresh
If .Text1.Text <> "" Then
cond2 = True
Else
cond2 = False
End If
End With
'«·‰Õﬁﬁ „‰ ÊÃÊœ „Ê—œÌ‰
With Frm_Mord
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from morden"
.Data1.Refresh
If .Text1.Text <> "" Then
cond3 = True
Else
cond3 = False
End If
End With

'„ﬁ«—‰… ⁄«„…
If cond1 = True And cond2 = True And cond3 = True Then
 Frm_Accounts.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
 Frm_Accounts.Data1.RecordSource = "select * from users where username='" & nowuser & "'"
 Frm_Accounts.Data1.Refresh
 If CBool(Frm_Accounts.Text5.Text) = True Then
  Command5.Enabled = True
  mnu_store.Enabled = True
 End If
Else
  Command5.Enabled = False
  mnu_store.Enabled = False
End If
 



With Frm_Clients
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from clients"
.Data1.Refresh
If .Text2.Text <> "" Then
 If CBool(Frm_Accounts.Text6.Text) = True Then
   Command12.Enabled = True
   mnu_clients_money.Enabled = True
   frm_daily.Command2.Enabled = True
 End If
Else
Command12.Enabled = False
mnu_clients_money.Enabled = False
frm_daily.Command2.Enabled = False
End If
End With

Unload Frm_Accounts
End Function

Private Sub userss_Click()
Frm_Accounts.Show
StayOnTop Frm_Accounts

End Sub

Public Function LoadSetting()
'Õœ «·√Ì«„ «·›«’·… ﬁ»·  «—ÌŒ Œ—»«‰ œÊ«¡
sharp = 0
End Function

Public Function Initial()
If disshow = True Then
' ›—Ì€ ÃœÊ· «·√œÊÌ… «·›«”œ…
With Frm_Disactive
Dim X As Integer
On Error Resume Next
.Data1.Recordset.MoveFirst
For X = 1 To .Data1.Recordset.RecordCount
.Data1.Recordset.Delete
.Data1.Recordset.MoveFirst
Next
End With

' ⁄»∆… «·ﬁ«⁄œ… »«·√œÊÌ… «·›«”œ… ›⁄·«

frm_store.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
frm_store.Data1.RecordSource = "select * from pharstore"
frm_store.Data1.Refresh
On Error Resume Next
frm_store.Data1.Recordset.MoveFirst

If frm_store.Text5.Text = "" Then
Exit Function
Else
 Dim I As Long
  ' Õ÷Ì— «·»—ÊÃ—” »«—
 FrmProgress.bar.Max = Data1.Recordset.RecordCount
 FrmProgress.bar.Min = 0
 FrmProgress.bar.Value = 0
 FrmProgress.Show
 StayOnTop FrmProgress
 For I = 1 To frm_store.Data1.Recordset.RecordCount
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
  FrmProgress.bar.Value = I
  frm_store.Data1.Recordset.MoveNext
 Next
 Unload FrmProgress
If Frm_Disactive.Data1.Recordset.RecordCount > Int(0) Then
 Frm_Disactive.Show
 StayOnTop Frm_Disactive
Else
Exit Function
End If
End If

End If

'«· ⁄·Ìﬁ« 
End Function
