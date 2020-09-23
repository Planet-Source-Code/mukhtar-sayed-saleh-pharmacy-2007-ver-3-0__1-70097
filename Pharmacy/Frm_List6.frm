VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Frm_List6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÞÇÆãÉ ÃÔßÇá ÇáÚÈæÉ"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "shapeof"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2.45745e5
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Pharmacy 3\pharmokhtar.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "shapes"
      RightToLeft     =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÜÜÜÜæÇÝÜÜÜÞ"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "Frm_List6.frx":0000
      Top             =   3000
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frm_List6.frx":0234
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "Frm_List6.frx":0248
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Frm_List6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
txtsearch.Text = Text1.Text

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hwnd

'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ
Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
Data1.RecordSource = "select * from shapes"
Data1.Refresh

End Sub
Private Sub Command1_Click()
Frm_AddDoa.Text7.Text = Text1.Text
selshape = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub txtsearch_Change()
If txtsearch.Text = "" Then Exit Sub

With Data1.Recordset

    .MoveFirst
                .FindFirst "[shapeof] like '" & txtsearch.Text & "*'"
    
    If .NoMatch = True Then
        MsgBox "ÇáÔßá ÛíÑ ãæÌæÏ", vbExclamation, "äÙÇã ÇáÕíÏáíøÉ 2007"
        txtsearch.Text = Empty
        Exit Sub
    End If
    
DBGrid1.Refresh
DBGrid1.ReBind

    
End With

End Sub

