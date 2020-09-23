VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frm_list4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÑÌÇÁ ÇÎÊíÇÑ ÇáÏæÇÁ"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_list4.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtsearch 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox cboSearch 
         Height          =   315
         ItemData        =   "frm_list4.frx":29C12
         Left            =   120
         List            =   "frm_list4.frx":29C1F
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Text            =   "ßæÏ ÇáÏæÇÁ"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "comname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2.45745e5
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "doname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   2.45745e5
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "peiceprice"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   2.45745e5
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "docount"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   2.45745e5
         Width           =   150
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1200
         OleObjectBlob   =   "frm_list4.frx":29C47
         Top             =   3960
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_list4.frx":29E7B
         Height          =   6135
         Left            =   120
         OleObjectBlob   =   "frm_list4.frx":29E8F
         TabIndex        =   8
         Top             =   1200
         Width           =   4815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "frm_list4.frx":2AA22
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   7560
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pharstore"
      RightToLeft     =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frm_list4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text <> "" Then
With frm_change
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
selcash2 = CDbl(Text3.Text)
.Text3.SetFocus
Unload Me
End With
Else
MsgBox "ÇáÑÌÇÁ ÊÍÏíÏ ÇáÏæÇÁ ÇáãØáæÈ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from pharstore"
Data1.Refresh
cboSearch.ListIndex = 2
End Sub

Private Sub txtsearch_Change()
If txtsearch.Text = "" Then Exit Sub

With Data1.Recordset

    .MoveFirst
    
    Select Case cboSearch.ListIndex
    
        'ÈÍË Úä ØÑíÞ ÇÓã ÇáÔÑßÉ
        Case 0
            .FindFirst "[comname] like '" & txtsearch.Text & "*'"
            
        'Úä ØÑíÞ ÇÓã ÇáÏæÇÁ
        Case 1
            .FindFirst "[doname] like '" & txtsearch.Text & "*'"
            
        'Úä ØÑíÞ ßæÏ ÇáÏæÇÁ
        Case 2
            .FindFirst "[docode] like '" & txtsearch.Text & "*'"
            
    End Select
    
    If .NoMatch = True Then
        MsgBox "ÇáÏæÇÁ ÛíÑ ãæÌæÏ", vbExclamation, "äÙÇã ÇáÕíÏáíøÉ 2007"
        txtsearch.Text = Empty
        Exit Sub
    End If
    
DBGrid1.Refresh
DBGrid1.ReBind

    
End With

End Sub

