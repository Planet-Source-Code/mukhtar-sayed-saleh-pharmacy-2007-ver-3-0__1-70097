VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frm_tip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "äÕíÍÉ Çáíæã"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Frm_tip.frx":0000
      Top             =   120
      Width           =   5055
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "Frm_tip.frx":000C
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÅÛáÇÞ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "äÕíÍÉ ÃÎÑì"
      Height          =   375
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_tip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "test.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Show it.
    Frm_tip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function




Private Sub Command1_Click()
    DoNextTip

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        label1.Text = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        label1.Text = Tips.Item(CurrentTip)
    End If
End Sub


