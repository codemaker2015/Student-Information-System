Attribute VB_Name = "TrayIcon"
Option Explicit

Type NOTIFYICONDATA
  cbSize              As Long
  hwnd                As Long
  uID                 As Long
  uFlags              As Long
  uCallbackMessage    As Long
  hIcon               As Long
  szTip               As String * 64
End Type
    
Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NULL = &H0

Declare Function Shell_NotifyIconA Lib "shell32" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Declare Function PostMessage Lib "user32" _
 Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long


