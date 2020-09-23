Attribute VB_Name = "modDLGMenuViewGP"
Option Explicit
'*** by nr  give some credits to this people
'The following changes were made to cCommonDialog.
'by   Steve McMahon
'21 Oct 2003 Ken Halter   Menu View
'***
' Thumbdlg sample from BlackBeltVB.com
' http://blackbeltvb.com
' Written by Matt Hart
' using "http://blackbeltvb.com/free/thumbdlg.htm"

Private Declare Function SendMessage _
   Lib "user32.dll" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long _
   , ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName _
   Lib "user32.dll" Alias "GetClassNameA" _
   (ByVal hwnd As Long, ByVal lpClassName As String _
   , ByVal nMaxCount As Long) As Long
Private Declare Function EnumChildWindows _
   Lib "user32.dll" (ByVal hWndParent As Long _
   , ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent _
   Lib "user32.dll" (ByVal hwnd As Long) As Long
 
Public Const WM_COMMAND As Long = &H111

Public Enum ViewStyles
   [_NoViewProcess] = 0
   vsLargeIcons = &H7029
   vsSmallIcons = &H702A
   vsList = &H702B
   vsDetails = &H702C
   '*** by nr this one I made up, why? because
   vsThumbnails = &H702D
End Enum
Private meLocalViewStyle As ViewStyles
Public meDesiredView As ViewStyles
Public Sub SetView(ByVal DLG As Long _
   , ByVal DesiredView As ViewStyles)
   Dim hwnd As Long
   meDesiredView = DesiredView
   'hwnd = GetParent(DLG)   'by nr
   hwnd = DLG
   Call EnumSelected(hwnd)
End Sub
 
Private Sub EnumSelected(hwnd As Long)
   Call EnumChildWindows(hwnd, AddressOf EnumChildProc, &H0)
End Sub
 
Public Function EnumChildProc(ByVal hwnd As Long, _
                              ByVal lParam As Long) As Long
 
   Dim lCont As Long
   Dim lRet As Long
   Dim sBuffer As String
    '***
    Dim tWnd As Long
   lCont = 1
 
   sBuffer = Space$(260)
   lRet = GetClassName(hwnd, sBuffer, Len(sBuffer))
 
   If InStr(1, sBuffer, "SHELLDLL_DefView") > 0 Then
   'by nr 'this works
      Call SendMessage(hwnd, WM_COMMAND, meDesiredView, ByVal 0&)
      lCont = 0 'stop looking for windows
'***
    End If

   EnumChildProc = lCont
 
End Function '=================================================================

