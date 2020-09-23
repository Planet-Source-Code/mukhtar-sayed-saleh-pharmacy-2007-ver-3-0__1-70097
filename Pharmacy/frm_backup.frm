VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáäÓÎ ÇáÅÍÊíÇØí"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_backup.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   6975
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command4 
         Caption         =   "ÇÓÊÚÑÇÖ ..."
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   4455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ÇÓÊÚÑÇÖ ..."
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "frm_backup.frx":29C12
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   160
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "frm_backup.frx":29C86
         TabIndex        =   10
         Top             =   675
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Width           =   6975
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÇÓÊÚÇÏÉ ÇáäÓÎÉ ÇáÇÍÊíÇØíøÉ"
         Height          =   855
         Left            =   3120
         Picture         =   "frm_backup.frx":29CFC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅäÔÇÁ äÓÎÉ ÇÍÊíÇØíøÉ"
         Height          =   855
         Left            =   5040
         Picture         =   "frm_backup.frx":2A479
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ÅÛáÇÞ"
         Height          =   855
         Left            =   0
         Picture         =   "frm_backup.frx":2ABFB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "frm_backup.frx":2B1F7
      Top             =   240
   End
End
Attribute VB_Name = "frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const OPEN_EXISTING = 3
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CopyFile& Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long)


Public C6Func As String


Private Sub Command1_Click()
If Text1.Text = "" Then
Text2.Text = Text2.Text & vbNewLine & "íÌÈ ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáåÏÝ ÃæáÇð"
Exit Sub
End If
Text2.Text = Text2.Text & vbNewLine & "íÊã ÇáÊÍÖíÑ áÚãáíøÉ ÇáäÓÎ ÇáÇÍÊíÇØí ÞÏ ÊÓÊÛÑÞ ÇáÚãáíøÉ ÈÖÚ ÏÞÇÆÞ"
On Error Resume Next
Text2.Text = Text2.Text & vbNewLine & "ÞØÚ ÇÊÕÇá ÞÇÚÏÉ ÇáÈíÇäÇÊ ãÚ ÇáäæÇÝÐ ÈÏÁ ÇáÂä"
'ÇÛáÇÞ ÌãíÚÚ ÇáäæÇÝÐ
Unload Frm_Accounts
Unload Frm_Add_Money
Unload Frm_AddClient
Unload Frm_AddDoa
Unload Frm_Addmord
Unload Frm_Addshape
Unload frm_cash
Unload frm_change
Unload Frm_client_cash
Unload Frm_Clients
Unload Frm_Clients_Find
Unload Frm_Clients_Money
Unload frm_companies
Unload frm_daily
Unload frm_date1
Unload frm_date2
Unload Frm_dAVE
Unload Frm_Disactive
Unload Frm_EndClient
Unload Frm_Increment
Unload frm_list
Unload frm_list2
Unload frm_list3
Unload frm_list4
Unload Frm_List5
Unload Frm_List6
Unload Frm_List7
Unload frm_login
Unload Frm_Mord
Unload Frm_Naf
Unload Frm_NewAccount
Unload frm_print
Unload frm_reports
Unload frm_saleend
Unload frm_SalePoint
Unload Frm_Shapes
Unload frm_store
Unload frm_store_find
Unload Frm_tip
Unload frmsetting
'ÞØÚ ÇáÇÊÕÇá ãÚ ÇáÏÇÊÇ ÇäÝÇíÑæäãíäÊÓ
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
Text2.Text = Text2.Text & vbNewLine & "ÞØÚ ÇÊÕÇá ÞÇÚÏÉ ÇáÈíÇäÇÊ ãÚ ÇáäæÇÝÐ ÇäÊåì ÇáÂä"
Text2.Text = Text2.Text & vbNewLine & "ÅáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ"
'ÇáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ æÇáßæãÇäÏÇÊ
With frm_main
.Command1.Enabled = False
.Command2.Enabled = False
.Command3.Enabled = False
.Command4.Enabled = False
.Command5.Enabled = False
.Command6.Enabled = False
.Command7.Enabled = False
.Command8.Enabled = False
.Command9.Enabled = False
.Command10.Enabled = False
.Command11.Enabled = False
.Command12.Enabled = False
.Command13.Enabled = False
.MnuFile.Enabled = False
.mnuhelp.Enabled = False
.mnu_tools.Enabled = False
.Mnu_setting.Enabled = False
End With
Text2.Text = Text2.Text & vbNewLine & "ÅáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ Êã ÈäÌÇÍ"
Text1.Enabled = False
Text3.Enabled = False
SkinLabel1.Enabled = False
SkinLabel2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Caption = "ÅáÛÇÁ ÇáÃãÑ"
C6Func = "ccopy"
'ÇáÂä ÊÍÖíÑ ÇáÈÑæÌÑÓ ÈÇÑ æ ÈÏÁ ÚãáíøÉ ÇáäÓÎ

    Dim hFile As Long, FileInfo As BY_HANDLE_FILE_INFORMATION
    hFile = CreateFile(App.Path & ("\pharmokhtar.dll"), 0, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    GetFileInformationByHandle hFile, FileInfo
    CloseHandle hFile
    FrmProgress.bar.Max = CLng(FileInfo.nFileSizeLow)
    FrmProgress.bar.Min = 0
    FrmProgress.Show
    'ãÊÛíÑ ãÔÇä ÇáÈÑæÌÑÓ ÈÇÑ
    Dim iscopy As Boolean
    iscopy = True
        
   CopyAFile = CopyFile(App.Path & ("\pharmokhtar.dll"), Text1.Text, Overwrite)
   
    Do While (iscopy = True)
    hFile = CreateFile(Text1.Text, 0, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    GetFileInformationByHandle hFile, FileInfo
    CloseHandle hFile
    FrmProgress.bar.Value = CLng(FileInfo.nFileSizeLow)
    DoEvents
    If FrmProgress.bar.Value = CLng(FileInfo.nFileSizeLow) Then
    iscopy = False
    Unload FrmProgress
    End If
    Loop

  If CopyAFile = 0 Then
    Text2.Text = Text2.Text & vbNewLine & "ÝÔáÊ  ÚãáíøÉ ÇäÔÇÁ ãáÝ ÇáäÓÎ ÇáÇÍÊíÇØí áÃÓÈÇÈ ÛíÑ ãÚÑæÝÉ !!"
    Exit Sub
  End If
Text2.Text = Text2.Text & vbNewLine & "Êã ÇäÔÇÁ ãáÝ ÇáäÓÎ  ÇáÇÍÊíÇØí ÈäÌÇÍ ÊÇã"
MsgBox "ÊãÊ ÚãáíøÉ ÇäÔÇÁ ÇáäÓÎ ÇáÇÍÊíÇØí ÈäÌÇÍ ÊÇã", vbInformation, "äÙÇã ÇáÕíÏáíøÉ 2007"
EndCopying
End Sub

Private Sub Command2_Click()
Dim mss
mss = MsgBox("ÊÍÐíÑ : ÓæÝ íÊã ÍÐÝ ßÇÝÉ ÇáÈíÇäÇÊ ÇáÍÇáíøÉ æÇáÚæÏÉ ááÈíÇäÇÊ ÇáãæÌæÏÉ Ýí ãáÝ ÇáäÓÎÉ ÇáÇÍÊíÇØíøÉ ÇáãÍÏÏÉ åá ÊÑíÏ ÇáÇÓÊãÑÇÑ ¿", vbInformation + vbYesNo, "äÙÇã ÇáÕíÏáíøÉ 2007")
If mss = vbNo Then
Exit Sub
End If

If Text3.Text = "" Then
Text2.Text = Text2.Text & vbNewLine & "íÌÈ ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáãÕÏÑ ÃæáÇð"
Exit Sub
End If

If FileExists(Text3.Text) = False Then
MsgBox "ÇáãáÝ ÇáãÕÏÑ ÛíÑ ãæÌæÏ Ãæ ÑÈãÇ íßæä ÊÇáÝÇð ÇáÑÌÇÁ ÇáÊÍÞÞ ãä ÇáÃãÑ æ ÇáãÍÇæáÉ ãä ÌÏíÏ", 16, "äÙÇã ÇáÕíÏáíøÉ 2007"
Exit Sub
End If

Text2.Text = Text2.Text & vbNewLine & "íÊã ÇáÊÍÖíÑ áÚãáíøÉ ÇáÇÓÊÚÇÏÉ ÞÏ ÊÓÊÛÑÞ ÇáÚãáíøÉ ÈÖÚ ÏÞÇÆÞ"
On Error Resume Next
Text2.Text = Text2.Text & vbNewLine & "ÞØÚ ÇÊÕÇá ÞÇÚÏÉ ÇáÈíÇäÇÊ ãÚ ÇáäæÇÝÐ ÈÏÁ ÇáÂä"
'ÇÛáÇÞ ÌãíÚÚ ÇáäæÇÝÐ
Unload Frm_Accounts
Unload Frm_Add_Money
Unload Frm_AddClient
Unload Frm_AddDoa
Unload Frm_Addmord
Unload Frm_Addshape
Unload frm_cash
Unload frm_change
Unload Frm_client_cash
Unload Frm_Clients
Unload Frm_Clients_Find
Unload Frm_Clients_Money
Unload frm_companies
Unload frm_daily
Unload frm_date1
Unload frm_date2
Unload Frm_dAVE
Unload Frm_Disactive
Unload Frm_EndClient
Unload Frm_Increment
Unload frm_list
Unload frm_list2
Unload frm_list3
Unload frm_list4
Unload Frm_List5
Unload Frm_List6
Unload Frm_List7
Unload frm_login
Unload Frm_Mord
Unload Frm_Naf
Unload Frm_NewAccount
Unload frm_print
Unload frm_reports
Unload frm_saleend
Unload frm_SalePoint
Unload Frm_Shapes
Unload frm_store
Unload frm_store_find
Unload Frm_tip
Unload frmsetting
'ÞØÚ ÇáÇÊÕÇá ãÚ ÇáÏÇÊÇ ÇäÝÇíÑæäãíäÊÓ
If DataEnvironment1.Connection1.State = 1 Then
DataEnvironment1.Connection1.Close
End If
Text2.Text = Text2.Text & vbNewLine & "ÞØÚ ÇÊÕÇá ÞÇÚÏÉ ÇáÈíÇäÇÊ ãÚ ÇáäæÇÝÐ ÇäÊåì ÇáÂä"
Text2.Text = Text2.Text & vbNewLine & "ÅáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ"
'ÇáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ æÇáßæãÇäÏÇÊ
With frm_main
.Command1.Enabled = False
.Command2.Enabled = False
.Command3.Enabled = False
.Command4.Enabled = False
.Command5.Enabled = False
.Command6.Enabled = False
.Command7.Enabled = False
.Command8.Enabled = False
.Command9.Enabled = False
.Command10.Enabled = False
.Command11.Enabled = False
.Command12.Enabled = False
.Command13.Enabled = False
.MnuFile.Enabled = False
.mnuhelp.Enabled = False
.mnu_tools.Enabled = False
.Mnu_setting.Enabled = False
End With
Text2.Text = Text2.Text & vbNewLine & "ÅáÛÇÁ ÊÝÚíá ÇáÃÒÑÇÑ Êã ÈäÌÇÍ"
Text1.Enabled = False
Text3.Enabled = False
SkinLabel1.Enabled = False
SkinLabel2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Caption = "ÅáÛÇÁ ÇáÃãÑ"
C6Func = "crestore"

'ÇáÂä ÊÍÖíÑ ÇáÈÑæÌÑÓ ÈÇÑ æ ÈÏÁ ÚãáíøÉ ÇáÇÓÊÚÇÏÉ

    Dim hFile As Long, FileInfo As BY_HANDLE_FILE_INFORMATION
    hFile = CreateFile(Text3.Text, 0, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    GetFileInformationByHandle hFile, FileInfo
    CloseHandle hFile
    FrmProgress.bar.Max = CLng(FileInfo.nFileSizeLow)
    FrmProgress.bar.Min = 0
    FrmProgress.Show
    'ãÊÛíÑ ãÔÇä ÇáÈÑæÌÑÓ ÈÇÑ
    Dim iscopy As Boolean
    iscopy = True
        
   CopyAFile = CopyFile(Text3.Text, App.Path & ("\pharmokhtar.dll"), Overwrite)
   
    Do While (iscopy = True)
    hFile = CreateFile(App.Path & ("\pharmokhtar.dll"), 0, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    GetFileInformationByHandle hFile, FileInfo
    CloseHandle hFile
    FrmProgress.bar.Value = CLng(FileInfo.nFileSizeLow)
    DoEvents
    If FrmProgress.bar.Value = CLng(FileInfo.nFileSizeLow) Then
    iscopy = False
    Unload FrmProgress
    End If
    Loop

  If CopyAFile = 0 Then
    Text2.Text = Text2.Text & vbNewLine & "ÝÔáÊ ÚãáíøÉ ÇáÇÓÊÚÇÏÉ Þã ÈÅÛáÇÞ Ãí ÈÑäÇãÌ ÂÎÑ íÓÊÚãá ÞÇÚÏÉ ÇáÈíÇäÇÊ Ëã ÍÇæá ãä ÌÏíÏ"
    Exit Sub
  End If
Text2.Text = Text2.Text & vbNewLine & "ÊãÊ ÚãáíøÉ ÇáÇÓÊÚÇÏÉ ÈäÌÇÍ ÊÇã"
MsgBox "ÊãÊ ÚãáíøÉ ÇáÇÓÊÚÇÏÉ ÈäÌÇÍ", vbInformation, "äÙÇã ÇáÕíÏáíøÉ 2007"
EndCopying

End Sub

Private Sub Command3_Click()
cd1.Filter = "ãáÝÇÊ äÓÎ ÇÍÊíÇØí áäÙÇã ÇáÕíÏáíøÉ |*.MBack"
cd1.DialogTitle = "ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáåÏÝ áÚãáíÉ ÇáäÓÎ ÇáÇÍÊíÇØí"
cd1.ShowSave
Text1.Text = cd1.FileName
If Text1.Text <> "" Then
Text2.Text = Text2.Text & vbNewLine & "Êã ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáåÏÝ"
End If
refreshcommands
End Sub

Private Sub Command4_Click()
cd1.Filter = "ãáÝÇÊ äÓÎ ÇÍÊíÇØí áäÙÇã ÇáÕíÏáíøÉ |*.MBack"
cd1.DialogTitle = "ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáãÕÏÑ áÚãáíÉ ÇáäÓÎ ÇáÇÍÊíÇØí"
cd1.ShowOpen
Text3.Text = cd1.FileName
If Text3.Text <> "" Then
Text2.Text = Text2.Text & vbNewLine & "Êã ÊÍÏíÏ ãÓÇÑ ÇáãáÝ ÇáãÕÏÑ"
End If
refreshcommands

End Sub

Private Sub Command6_Click()
If C6Func = "close" Then
Unload Me
End If
If C6Func = "ccopy" Then
EndCopying
End If
If C6Func = "crestore" Then
EndCopying
End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\mokhtar.skn")
Skin1.ApplySkin Me.hWnd
refreshcommands
If FileExists(App.Path & ("\pharmokhtar.dll")) = True Then
Text2.Text = "Êã ÇáÇÊÕÇá ÈãáÝ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáÍÇáíøÉ ÈäÌÇÍ"
Else
Text2.Text = "áÇ íãßä ÊÚííä ãÓÇÑ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáÍÇáíøÉ"
End If
C6Func = "close"
End Sub


Public Function refreshcommands()
If Text1.Text = "" And Text3.Text = "" Then
Command1.Enabled = False
Command2.Enabled = False
End If
If Text1.Text <> "" And Text3.Text = "" Then
Command1.Enabled = True
Command2.Enabled = False
End If
If Text1.Text = "" And Text3.Text <> "" Then
Command1.Enabled = False
Command2.Enabled = True
End If
If Text1.Text <> "" And Text3.Text <> "" Then
Command1.Enabled = True
Command2.Enabled = True
End If
End Function


Public Function EndCopying()
With frm_main
.Command1.Enabled = True
.Command2.Enabled = True
.Command3.Enabled = True
.Command4.Enabled = True
.Command5.Enabled = True
.Command6.Enabled = True
.Command7.Enabled = True
.Command8.Enabled = True
.Command9.Enabled = True
.Command10.Enabled = True
.Command11.Enabled = True
.Command12.Enabled = True
.Command13.Enabled = True
.MnuFile.Enabled = True
.mnuhelp.Enabled = True
.mnu_tools.Enabled = True
.Mnu_setting.Enabled = True
End With
Text1.Enabled = True
Text3.Enabled = True
SkinLabel1.Enabled = True
SkinLabel2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Caption = "ÅÛáÇÞ"
C6Func = "close"
With frm_main
.Refreshcommand
End With
End Function

