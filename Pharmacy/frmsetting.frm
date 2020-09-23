VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmsetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÚÏÇÏÇÊ ÇáäÙÇã"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
      DataField       =   "pharpriceshow"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Text            =   "Text19"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Alignment       =   1  'Right Justify
      DataField       =   "disshow"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Text            =   "Text18"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      DataField       =   "tip"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Text            =   "Text17"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      DataField       =   "dayssharp"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Text            =   "Text16"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pharsets"
      RightToLeft     =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      DataField       =   "phone1"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text10"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      DataField       =   "city"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text14"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      DataField       =   "mailbox"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text13"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      DataField       =   "fax"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text12"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      DataField       =   "phone2"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text11"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Frame Frame2 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ÅÛáÇÞ"
            Height          =   735
            Left            =   120
            Picture         =   "frmsetting.frx":0000
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   5400
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Caption         =   "ÅÚÏÇÏÇÊ ÇáäÙÇã"
            Height          =   975
            Left            =   0
            Picture         =   "frmsetting.frx":05FC
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Caption         =   "ãÚáæãÇÊ ÇáÕíÏáíøÉ"
            Height          =   975
            Left            =   0
            Picture         =   "frmsetting.frx":0DB0
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Frame Frame6 
            Height          =   2415
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   120
            Width           =   6015
            Begin VB.CommandButton Command2 
               Caption         =   "ãÜæÇÝÜÞ"
               Height          =   375
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "                             ÅÙåÇÑ äÕíÍÉ Çáíæã ÚäÏ ßá ÊÔÛíá ááäÙÇã"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   960
               Width           =   4335
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   "  ÅÙåÇÑ ÇáÃÏæíÉ ÐÇÊ ÇáÕáÇÍíøÉ ÇáãäÊåíÉ Ýí ÈÏÇíÉ ÊÔÛíá ÇáäÙÇã"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1440
               Width           =   4335
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               Caption         =   "                                   ÅÙåÇÑ ÓÚÑ ÇáÕíÏáí Ýí ÔÇÔÉ ÇáÈíÚ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1920
               Visible         =   0   'False
               Width           =   4335
            End
            Begin VB.TextBox Text15 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   480
               Left            =   2040
               MaxLength       =   2
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Text            =   "0"
               Top             =   240
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   3600
               OleObjectBlob   =   "frmsetting.frx":1634
               TabIndex        =   33
               Top             =   360
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   1560
               OleObjectBlob   =   "frmsetting.frx":16D4
               TabIndex        =   34
               Top             =   360
               Width           =   375
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Frame Frame7 
            Height          =   4335
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   6015
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   100
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   240
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   255
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   25
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   1200
               Width           =   4455
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   25
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   1680
               Width           =   4455
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   25
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   2160
               Width           =   4455
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   2640
               Width           =   4455
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   3120
               Width           =   4455
            End
            Begin VB.CommandButton Command1 
               Caption         =   "ãæÇÝÞ"
               Height          =   375
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   3840
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":1738
               TabIndex        =   24
               Top             =   360
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":17B0
               TabIndex        =   25
               Top             =   840
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":182C
               TabIndex        =   26
               Top             =   1320
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":18AC
               TabIndex        =   27
               Top             =   1800
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":192C
               TabIndex        =   28
               Top             =   2280
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":1992
               TabIndex        =   29
               Top             =   2760
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "frmsetting.frx":1A04
               TabIndex        =   30
               Top             =   3240
               Width           =   1215
            End
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            DataField       =   "pharadmin"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Text            =   "Text9"
            Top             =   2.45745e5
            Width           =   495
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            DataField       =   "pharname"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Text            =   "Text8"
            Top             =   2.45745e5
            Width           =   855
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   1920
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "pharinformation"
            RightToLeft     =   -1  'True
            Top             =   4320
            Visible         =   0   'False
            Width           =   1740
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   495
            Left            =   120
            OleObjectBlob   =   "frmsetting.frx":1A70
            TabIndex        =   5
            Top             =   5520
            Width           =   6015
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "frmsetting.frx":1B45
      Top             =   5640
   End
End
Attribute VB_Name = "frmsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function refreshframes()
If Option1.Value = True Then
Frame3.Visible = True
Else
Frame3.Visible = False
End If

If Option2.Value = True Then
Frame4.Visible = True
Else
Frame4.Visible = False
End If



End Function

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇÓã ÇáÕíÏáíøÉ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If Text8.Text <> "" Then
  Data1.Recordset.Edit
  Text8.Text = CStr(Text1.Text)
  Text9.Text = CStr(Text2.Text)
  Text10.Text = CStr(Text3.Text)
  Text11.Text = CStr(Text4.Text)
  Text12.Text = CStr(Text5.Text)
  Text13.Text = CStr(Text6.Text)
  Text14.Text = CStr(Text7.Text)
  On Error Resume Next
  Data1.Recordset.MoveNext
  Data1.Recordset.MovePrevious
Else
  Data1.Recordset.AddNew
  Text8.Text = CStr(Text1.Text)
  Text9.Text = CStr(Text2.Text)
  Text10.Text = CStr(Text3.Text)
  Text11.Text = CStr(Text4.Text)
  Text12.Text = CStr(Text5.Text)
  Text13.Text = CStr(Text6.Text)
  Text14.Text = CStr(Text7.Text)
  On Error Resume Next
  Data1.Recordset.MoveNext
  Data1.Recordset.MovePrevious

End If
readsetting (False)

End Sub

Private Sub Command2_Click()
If Text15.Text = "" Then
MsgBox "ÇáÑÌÇÁ ÅÏÎÇá ÇÏÎÇá ÝÑÞ ÇáÃíøÇã ÞÈá ÇÚÊÈÇÑ ÇáÏæÇÁ ÝÇÓÏ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If IsNumeric(Text15.Text) = False Then
MsgBox "Åä ÇáÊÚÈíÑ ÇáÐí Ýí ÝÑÞ ÇáÃíÇã ÞÈá ÇÚÊÈÇÑ ÇáÏæÇÁ ÝÇÓÏ áíÓ ÑÞãÇð ÇáÑÌÇÁ ÇáÊÍÞÞ ãäå", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If CInt(Text15.Text) < CInt(0) Then
MsgBox "áÇ íÌæÒ Ãä íßæä ÝÑÞ ÇáÃíÇã ÞÈá ÇÚÊÈÇÑ ÇáÏæÇÁ ÝÇÓÏ ÚÏÏ ÓÇáÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

If Text16.Text <> "" Then
  Data2.Recordset.Edit
  Text16.Text = CInt(Text15.Text)
  Text17.Text = CBool(Check1.Value)
  Text18.Text = CBool(Check2.Value)
  Text19.Text = CBool(Check3.Value)
  On Error Resume Next
  Data2.Recordset.MoveNext
  Data2.Recordset.MovePrevious
Else
  Data2.Recordset.AddNew
  Text16.Text = CInt(Text15.Text)
  Text17.Text = CBool(Check1.Value)
  Text18.Text = CBool(Check2.Value)
  Text19.Text = CBool(Check3.Value)
  On Error Resume Next
  Data2.Recordset.MoveNext
  Data2.Recordset.MovePrevious
End If

readsetting (False)

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd
'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharinformation"
Data1.Refresh
' ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáËÇäíÉ
Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data2.RecordSource = "select * from pharsets"
Data2.Refresh

Me.refreshframes
Me.Refreshvalue
End Sub

Private Sub Option1_Click()
Me.refreshframes
End Sub

Private Sub Option2_Click()
Me.refreshframes

End Sub

Private Sub Option3_Click()
Me.refreshframes
End Sub


Public Function Refreshvalue()
  Text1.Text = CStr(Text8.Text)
  Text2.Text = CStr(Text9.Text)
  Text3.Text = CStr(Text10.Text)
  Text4.Text = CStr(Text11.Text)
  Text5.Text = CStr(Text12.Text)
  Text6.Text = CStr(Text13.Text)
  Text7.Text = CStr(Text14.Text)
  Text15.Text = CStr(Text16.Text)
  On Error Resume Next
  If CBool(Text17.Text) = True Then
  Check1.Value = 1
  Else
  Check1.Value = 0
  End If
  
  If CBool(Text18.Text) = True Then
  Check2.Value = 1
  Else
  Check2.Value = 0
  End If

  If CBool(Text19.Text) = True Then
  Check3.Value = 1
  Else
  Check3.Value = 0
  End If

End Function

