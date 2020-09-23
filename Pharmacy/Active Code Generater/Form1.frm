VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pharmacy Active Code Creater"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate My Active Code"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   6495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Text            =   "31032156465461321564987"
      Top             =   1560
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter_Client_Code_Here"
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
getactive (Text1.Text)

End Sub

Private Sub Command2_Click()
End

End Sub

Public Function getactive(ByVal sn As Long)
On Error Resume Next
Dim ac As String
Dim X, Y, z As Double
Dim logx, logy, logz As Double
Dim res1 As Double
X = Mid(sn, 1, 2)
Y = Mid(sn, 3, 6)
z = Right(sn, 3)
logx = Log(X)
logy = Log(Y)
logz = Log(z)
res1 = CLng(logx) * CLng(logy) * CLng(logz) * (CLng(logx) / 2) + 1
ac = (res1 * 8254) + 4445
Text2.Text = ac
End Function

