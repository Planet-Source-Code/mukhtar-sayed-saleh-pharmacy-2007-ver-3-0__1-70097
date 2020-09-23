VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form Frm_List5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÞÇÆãÉ ÇáÚãáÇÁ"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_List5.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "clientcode"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "clientname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   735
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÜÜÜÜæÇÝÜÜÜÞ"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clients"
      RightToLeft     =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "Frm_List5.frx":29C12
      Top             =   3120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   3240
      OleObjectBlob   =   "Frm_List5.frx":29E46
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_List5.frx":29EB8
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "Frm_List5.frx":29ECC
      TabIndex        =   4
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "Frm_List5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Frm_client_cash.Text1.Text = Text1.Text
selcode2 = Text2.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from clients"
Data1.Refresh

End Sub

Private Sub txtsearch_Change()
If txtsearch.Text = "" Then Exit Sub

With Data1.Recordset

    .MoveFirst
                .FindFirst "[clientname] like '" & txtsearch.Text & "*'"
    
    If .NoMatch = True Then
        MsgBox "ÇáÚãíá ÛíÑ ãæÌæÏ", vbExclamation, "äÙÇã ÇáÕíÏáíøÉ 2007"
        txtsearch.Text = Empty
        Exit Sub
    End If
    
DBGrid1.Refresh
DBGrid1.ReBind

    
End With

End Sub

