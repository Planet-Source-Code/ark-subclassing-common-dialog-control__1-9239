Attribute VB_Name = "Module1"
Type POINTAPI
     x As Long
     y As Long
End Type

Type RECT
     left As Long
     top As Long
     right As Long
     bottom As Long
End Type

Enum ActionType ' Notification massages from controls
     EN_SETFOCUS = &H100   ' TextBox Receive Focus
     EN_KILLFOCUS = &H200  ' TextBox Lost Focus
     EN_CHANGE = &H300     ' After Text in TextBox Change
     EN_UPDATE = &H400     ' Before Text in TextBox Change
     CBN_SELCHANGE = 1     ' Change Selection in Combo
     CBN_SELENDCANCEL = 10 ' Close combo without selection
     CBN_SELENDOK = 9      ' Close combo with selection
     CBN_KILLFOCUS = 4     ' Combo lost focus
     CBN_CLOSEUP = 8       ' Close DropDown List
     CBN_DROPDOWN = 7      ' Open DropDown List
     CBN_SETFOCUS = 3      ' Combo receive focus
     BN_CLICKED = 0        ' Button was clicked
End Enum

Enum CtrlID ' ID of controls
' For All dialogs
     ID_OK = &H1  'Open or Save button
     ID_CANCEL = &H2 'Cancel Button
     ID_HELP = &H40E 'Help Button
' For open/save dialogs
     ID_READONLY = &H410 'Read-only check box
     ID_FILETYPELABEL = &H441 'FileType label
     ID_FILELABEL = &H442 'FileName label
     ID_FOLDERLABEL = &H443 'Folder label
     ID_LIST = &H461 'Parent of file list
     ID_FILETYPE = &H470 'FileType combo box
     ID_FOLDER = &H471 'Folder combo box
     ID_FILETEXT = &H480 'FileName text box
     ID_NEWFOLDER = &HFFFFA002 ' NewFolder Button - can not be disabled
     ID_PARENTFOLDER = &HFFFFA001  ' GoUp button  - can not be disabled
' for print dialogs
     ID_ALLPAGES = &H420
     ID_SELECTEDTEXT = &H421
     ID_PAGERANGE = &H422
     ID_COPYES = &H482
     ID_PRINTERCOMBO = &H473
     ID_PRINTTOFILE = &H410
     ID_PROPERTIES = &H401
 ' for font Dialog
     ID_FONTTEXT = &H470
     ID_STYLETEXT = &H471
     ID_SIZETEXT = &H471
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Const WM_SETFOCUS = &H7
Const WM_CLOSE = &H10
Const WM_COMMAND = &H111
Const WM_DESTROY = &H2
Const WM_ENABLE = &HA
Const GWL_WNDPROC = (-4)
Const GWL_STYLE = (-16)
Const WS_DISABLED = &H8000000
Const NV_DLG As Long = &H5000&
Const WM_CONTEXTMENU = &H7B

Dim OldProcess As Long, f As Form, hDlg As Long, hDlgMenu As Long, OldFileList As Long

Public Sub SetControlOnDlg(fOwner As Form)
  Set f = fOwner
  SetTimer f.hwnd, NV_DLG, 0&, AddressOf TimerProc
End Sub

Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
  hDlg = FindWindow("#32770", vbNullString)
  KillTimer hwnd, idEvent
  f.Cdlg_Init
  OldProcess = GetWindowLong(hDlg, GWL_WNDPROC)
  SetWindowLong hDlg, GWL_WNDPROC, AddressOf DlgProc
  Dim hFileList As Long
  hFileList = GetDlgItem(hDlg, ID_LIST)
  OldFileList = GetWindowLong(hFileList, GWL_WNDPROC)
  SetWindowLong hFileList, GWL_WNDPROC, AddressOf FileListProc
End Sub

Public Function FileListProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Static bTemp As Boolean
  If wMsg = WM_CONTEXTMENU Then
     bTemp = Not bTemp
     If bTemp Then MsgBox "PopUp menu is disabled"
  Else
     FileListProc = CallWindowProc(OldFileList, hwnd, wMsg, wParam, lParam)
  End If
End Function

Public Function DlgProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case wMsg
     Case WM_DESTROY
          SetWindowLong hwnd, GWL_WNDPROC, OldProcess
          Set f = Nothing
     Case WM_CLOSE
     Case WM_COMMAND
          Dim id As Long, bCancel As Boolean
          id = GetLoWord(wParam)
          f.Cdlg_UserAction id, bCancel, GetHiWord(wParam)
          If bCancel Then Exit Function
  End Select
  DlgProc = DefDlgProc(hwnd, wMsg, wParam, lParam)
End Function

Public Sub ModifyCtrl(nItem As CtrlID, Optional sNewText As String = "", Optional bEnabled As Boolean = True, Optional bVisible As Boolean = True)
   Dim hItem As Long
   hItem = GetDlgItem(hDlg, nItem)
   If sNewText <> "" Then SetWindowText hItem, sNewText
   If Not bEnabled Then
      SetWindowLong hItem, GWL_STYLE, GetWindowLong(hItem, GWL_STYLE) Or WS_DISABLED
      SendMessage hItem, WM_ENABLE, 0&, 0&
      SendMessage hDlg, WM_SETFOCUS, hItem, 0&
   End If
   If bVisible = False Then MoveWindow hItem, 0, 0, 0, 0, 1
End Sub

Public Sub MoveCtrl(nItem As CtrlID, Optional lLeft As Long = 0, Optional lTop As Long = 0, Optional lWidth As Long = 0, Optional lHeight As Long = 0)
   Dim hItem As Long, rc As RECT, pt As POINTAPI
   hItem = GetDlgItem(hDlg, nItem)
   GetWindowRect hItem, rc
   If lLeft = 0 Then lLeft = rc.left
   If lTop = 0 Then lTop = rc.top
   If lWidth = 0 Then lWidth = rc.right - rc.left
   If lHeight = 0 Then lHeight = rc.bottom - rc.top
   pt.x = lLeft
   pt.y = lTop
   ScreenToClient hDlg, pt
   MoveWindow hItem, pt.x, pt.y, lWidth, lHeight, 1&
End Sub

Public Sub MoveDialog(Optional lLeft As Long = 0, Optional lTop As Long = 0, Optional lWidth As Long = 0, Optional lHeight As Long = 0)
  Dim rc As RECT
  GetWindowRect hDlg, rc
  If lLeft = 0 Then lLeft = rc.left
  If lTop = 0 Then lTop = rc.top
  If lWidth = 0 Then lWidth = rc.right - rc.left
  If lHeight = 0 Then lHeight = rc.bottom - rc.top
  MoveWindow hDlg, lLeft, lTop, lWidth, lHeight, 1&
End Sub

Public Sub CenterDialog()
  Dim rc As RECT
  GetWindowRect hDlg, rc
  lLeft = (Screen.Width / Screen.TwipsPerPixelX - rc.right + rc.left) / 2
  lTop = (Screen.Height / Screen.TwipsPerPixelY - rc.bottom + rc.top) / 2
  MoveWindow hDlg, lLeft, lTop, rc.right - rc.left, rc.bottom - rc.top, 1&
End Sub

Public Function GetCtrlText(id As CtrlID) As String
   Dim sText As String, k As Long, hItem As Long
   hItem = GetDlgItem(hDlg, id)
   sText = Space$(512)
   k = GetWindowText(hItem, sText, 512)
   If k > 0 Then sText = left$(sText, k)
   GetCtrlText = sText
End Function

Private Function GetHiWord(dw As Long) As Long
  If dw And &H80000000 Then
     GetHiWord = (dw \ 65535) - 1
  Else
     GetHiWord = dw \ 65535
  End If
End Function

Private Function GetLoWord(dw As Long) As Long
   If dw And &H8000& Then
      GetLoWord = &H8000 Or (dw And &H7FFF&)
   Else
      GetLoWord = dw And &HFFFF&
   End If
End Function


