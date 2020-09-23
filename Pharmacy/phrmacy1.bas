Attribute VB_Name = "phrmacy1"
Option Explicit
Const WM_DISPLAYCHANGE = &H7E
Const HWND_BROADCAST = &HFFFF&
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const BITSPIXEL = 12
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public OldX As Long, OldY As Long, nDC As Long
Public nowuser As String
Public selected As String
Public selcash As Double
Public selcount As Integer
Public selcash2 As Double
Public selcode As Long
Public selcode2 As Long
Public selname2 As String
Public selshape As String
Public mokcommand As String
Public selmord As String
Public summ1 As Double
Public storecommand1 As String
Public storecommand2 As String
Public storecommand3 As String
Public pointcommand As String
Public mordcommand1 As String
Public shapecommand1 As String
Public sharp As Integer
Public tip As Boolean
Public disshow As Boolean
Public pharprice As Boolean
Public bcodenew As String
Public mypharname As String
Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2


'ÏÇáÉÏæãÇð Ýí ÇáãÞÏãÉ
Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Rem ÏÇáÉ ááÊÍÞÞ ãä æÌæÏ ãáÝ ãÚíøä
Public Function FileExists(strPath As String) As Boolean
    strPath = Trim(strPath)
    If strPath = "" Then
        FileExists = False
        Exit Function
    End If
  FileExists = Len(Dir(strPath)) <> 0
End Function

'ÏÇáÉ ãÎÊÇÑíøÉ áÞÑÇÁÉ ÅÚÏÇÏÇÊ ÇáÈÑäÇãÌ Ýí ÇÑÈÚÉ ãÊÞíÑÇÊ ÚÇãøÉ
Public Function readsetting(ByVal finalexit As Boolean)
With frmsetting
' ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáËÇäíÉ
.Data2.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data2.RecordSource = "select * from pharsets"
.Data2.Refresh
'ÊÍãíá ÇáÅÚÏÇÏÇÊ Ýí ÇáãÊÛíÑÇÊ
On Error Resume Next
  sharp = CInt(.Text16.Text)
  tip = CBool(.Text17.Text)
  disshow = CBool(.Text18.Text)
  pharprice = CBool(.Text19.Text)
 'ÝÊÍ ÞÇÚÏÉ ÇáÈíÇäÇÊ ÇáÃæáì
.Data1.DatabaseName = App.Path & ("\pharmokhtar.dll")
.Data1.RecordSource = "select * from pharinformation"
.Data1.Refresh
 
mypharname = .Text8.Text
End With
If finalexit = True Then
On Error Resume Next
Unload frmsetting
End If
End Function


Public Function bcodegen(ByVal N As Long)
'ÊæáíÏ ßæÏ ãä 12 ÎÇäÉ
Dim num As Long, str As String

Randomize
num = (1000000000 * Rnd) + 1 + CLng(N)
str = CStr(num) & CStr(1548668)
bcodenew = Mid(CStr(str), 1, 12)
If Len(bcodenew) < 12 Then
bcodenew = CStr(bcodenew) & CStr(0)
End If
End Function
