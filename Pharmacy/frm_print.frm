VERSION 5.00
Begin VB.Form frm_print 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   Picture         =   "frm_print.frx":0000
   ScaleHeight     =   3465
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command2 
         Caption         =   "ÅáÛÇÁ ÇáÃãÑ"
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ØÈÇÚÉ"
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   7095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   7095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   7095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frm_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_print.Command1.Value = False
frm_print.Command2.Value = False
frm_print.PrintForm
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
