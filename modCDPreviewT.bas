Attribute VB_Name = "modCDPreviewGP"
' ********************************************************************************
'
'  Module:  modCommonDialog
'  Author:  Garrett Sever   (garrett@elitevb.com)
'
' ********************************************************************************
'
'  Description: Module for customizing the common dialog's Open dialog with a
'               preview window when the user clicks on BMP, GIF, or JPG items.
'
' ********************************************************************************
'      Visit http://www.elitevb.com for more high-powered solutions!!
' ********************************************************************************
'*** by nr, I removed some declarations not used here
Option Explicit
'Dim oIPS As IPersistStream
Private moPvwImg            As StdPicture   ' Picture we display in the preview window
Private msPvwImgPath        As String       ' Path to the file for our previewed picture
Private mbUsePreview        As Boolean      ' Whether or not we display the preview window

' A point UDT
Private Type POINTAPI
    x As Long
    y As Long
End Type

' A rectangle UDT
Private Type RECT
    left    As Long
    top     As Long
    right   As Long
    bottom  As Long
End Type
'*** by nr  for View Menu
'Private Const WM_CREATE = &H1
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP As Long = &H4 '  left button up
'Const WS_CHILD = &H40000000
'Private Const WM_LBUTTONDOWN As Long = &H201
'Private Const WM_LBUTTONUP As Long = &H202
'Const SW_HIDE = 0
'Const SW_NORMAL = 1

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

'Public hHook As Long
Public hToolbar As Long

'***
' Used with the SetWindowPlacement API function
Private Type WINDOWPLACEMENT
    length              As Long
    flags               As Long
    showCmd             As Long
    ptMinPosition       As POINTAPI
    ptMaxPosition       As POINTAPI
    rcNormalPosition    As RECT
End Type

' OPENFILENAME structure. See MSDN for a complete explanation of all values
Private Type OPENFILENAME
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

' Bitmap structure - used to get width and height info
Private Type BITMAP
    bmType              As Long
    bmWidth             As Long
    bmHeight            As Long
    bmWidthBytes        As Long
    bmPlanes            As Integer
    bmBitsPixel         As Integer
    bmBits              As Long
End Type

' WM_PAINT structure - used for both the preview window and the preview info window.
Private Type PAINTSTRUCT
    hdc                 As Long
    fErase              As Long
    rcPaint             As RECT
    fRestore            As Long
    fIncUpdate          As Long
    rgbReserved         As Byte
End Type

' WM_NOTIFY Header structure
Private Type NMHDR
    hwndFrom            As Long       ' Window handle of control sending message
    idFrom              As Long       ' Identifier of control sending message
    code                As Long       ' Specifies the notification code
End Type

' OPENFILENAME common dialog flag constants
Private Const OFN_ENABLEHOOK    As Long = &H20
Private Const OFN_EXPLORER      As Long = &H80000
Private Const OFN_HIDEREADONLY  As Long = &H4

' Common Windows messages
Private Const WM_DESTROY        As Long = &H2
Private Const WM_PAINT          As Long = &HF
Private Const WM_NOTIFY         As Long = &H4E
Private Const WM_INITDIALOG     As Long = &H110
Private Const WM_USER           As Long = &H400

' Constants specific to the common dialog - ListView HWND and
'  two notification messages.
Private Const ID_LIST           As Long = &H460
Private Const CDN_INITDONE      As Long = -601      ' initialize dialog stuff is done
Private Const CDN_SELCHANGE     As Long = -602      ' When the selection is changed

' Constants used for showing our API created controls
Private Const SW_NORMAL         As Long = 1
Private Const SW_SHOWNORMAL     As Long = 1
' Constants used in CreateWindowEx
Private Const WS_EX_CLIENTEDGE  As Long = &H200
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Private Const WS_CHILD          As Long = &H40000000
Private Const WS_GROUP          As Long = &H20000
Private Const WS_VISIBLE        As Long = &H10000000

' More common dialog messages
Private Const CDM_FIRST = WM_USER + 100
Private Const CDM_GETFILEPATH = CDM_FIRST + &H1

' DrawText API constants
Private Const DT_BOTTOM                 As Long = &H8
Private Const DT_CALCRECT               As Long = &H400
Private Const DT_CENTER                 As Long = &H1
Private Const DT_EDITCONTROL            As Long = &H2000
Private Const DT_END_ELLIPSIS           As Long = &H8000
Private Const DT_EXPANDTABS             As Long = &H40
Private Const DT_EXTERNALLEADING        As Long = &H200
Private Const DT_HIDEPREFIX             As Long = &H100000
Private Const DT_INTERNAL               As Long = &H1000
Private Const DT_LEFT                   As Long = &H0
Private Const DT_MULTILINE              As Long = (&H1)
Private Const DT_NOCLIP                 As Long = &H100
Private Const DT_NOFULLWIDTHCHARBREAK   As Long = &H80000
Private Const DT_NOPREFIX               As Long = &H800
Private Const DT_PATH_ELLIPSIS          As Long = &H4000
Private Const DT_PREFIXONLY             As Long = &H200000
Private Const DT_RIGHT                  As Long = &H2
Private Const DT_RTLREADING             As Long = &H20000
Private Const DT_SINGLELINE             As Long = &H20
Private Const DT_TABSTOP                As Long = &H80
Private Const DT_TOP                    As Long = &H0
Private Const DT_VCENTER                As Long = &H4
Private Const DT_WORD_ELLIPSIS          As Long = &H40000
Private Const DT_WORDBREAK              As Long = &H10

' GetWindowLong and SetWindowLong constants
Private Const GWL_WNDPROC               As Long = -4

' API Declarations to make all this magic happen ;-)
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Declares for storing info against the window handles
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'*** by nr
Private hWndParent As Long
Public Declare Function SetWindowText Lib "user32" _
   Alias "SetWindowTextA" _
  (ByVal hwnd As Long, _
   ByVal lpString As String) As Long

'***

'*********************************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************
' OPENFILENAME stuff
'*********************************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************

Public Function OpenPictureDialog(aHwnd As Long, Optional ByVal ShowPreview As Boolean) As String

    Dim OFName As OPENFILENAME
    Dim sTemp As String     'by N, to eleiminate 0
    Static sInitDir As String

    ' Make sure that all of our subclassing and callback functions don't have any
    '  typos before we start subclassing. This helps prevent crashes after changes.
    OFNCallbackProc 0, 0, 0, 0
    PreviewWndProc 0, 0, 0, 0
    PreviewInfoWndProc 0, 0, 0, 0
    SetPreview 0
    
    
    mbUsePreview = ShowPreview
'***  by N
    If (sInitDir = "") Then
        sInitDir = App.Path
    End If
'***

    ' Populate the necessary values in our OFName UDT
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = aHwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Gif animation files (*.gif,*.gjm)" & Chr$(0) & "*.gif;*.gjm" & Chr$(0) & _
                         "GIF files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & _
                         "GJM files (*.gjm)" + Chr$(0) & "*.gjm" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = sInitDir
    OFName.lpstrTitle = "Open File with Preview"
    OFName.flags = OFN_ENABLEHOOK Or OFN_EXPLORER Or OFN_HIDEREADONLY
    OFName.lpfnHook = AddressOfFunction(AddressOf OFNCallbackProc)
    
    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        sTemp = Trim$(OFName.lpstrFile) 'by N
        If (Asc(Mid(sTemp, Len(sTemp), 1))) = 0 Then
            sTemp = Mid(sTemp, 1, Len(sTemp) - 1)
            sInitDir = ParsePath(sTemp, PATH_ONLY)
            OpenPictureDialog = sTemp
        Else
            OpenPictureDialog = sTemp
        End If
        'OpenPictureDialog = Trim$(OFName.lpstrFile) 'return asci 0 at the end
    Else
        OpenPictureDialog = ""
    End If
End Function

Public Function AddressOfFunction(aLongAddr As Long) As Long
    ' Cheap hack for using the AddressOf operator in the UDT
    AddressOfFunction = aLongAddr
End Function

Public Function OFNCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim anOFNotify  As NMHDR            ' a WM_NOTIFY message header
    Dim aStrAr()    As Byte             ' a byte array used for getting the selected object path
    Dim aStr        As String           ' a string variable (converted from aStrAr)
    Dim strLen      As Long             ' Length of the string
    Dim aWindPl     As WINDOWPLACEMENT  ' A WINDOWPLACEMENT UDT used to resize the dialog box
    Dim ptList      As POINTAPI         ' The upper left corner of the listview
    Dim ptWindow    As POINTAPI         ' the upper left corner of the dialog window
    Dim rcList      As RECT             ' The bounding rectangle for the listview
    Dim lstHwnd     As Long             ' Handle to the listview
    Dim lPvwLbl     As Long             ' Handle to the preview label we create
    Dim lPvwArea    As Long             ' Handle to the preview area where we draw
    Dim lPvwInfo    As Long             ' Handle to the dimension label for the picture
    
    If hwnd = 0 Then Exit Function
    
    On Local Error Resume Next
    Select Case uMsg
        Case WM_INITDIALOG
        ' Do whatever we want to when the dialog first
        'initializes. Consider this your Form_Load event ;-)

'*** by nr set View Menu to Thumbnails
        'this doesn't works here ?
        'Call SetView(GetParent(hwnd), vsDetails)

        'obtain the handle to the parent dialog
         hWndParent = GetParent(hwnd)
         
         If hWndParent <> 0 Then
         
           'Just to prove the handle was obtained,
           'change the dialog's caption
            Call SetWindowText(hWndParent, "I'm Hooked on Hooked Dialogs!")
            
            'by nr this works fine here
            'parent->PostMessage(WM_COMMAND, 40964, NULL);  'details
'''            Call PostMessage(hWndParent, WM_COMMAND, 40964, vbNull)
            '40962  create new folder
            '40961  up one level, same as 40967
            '40965  open "Look in"
        End If
'***
        Case WM_NOTIFY
            ' Copy the info from our pointer into a structure
            CopyMemory anOFNotify, ByVal lParam, Len(anOFNotify)
            If (anOFNotify.code = CDN_INITDONE) Then
            
                If mbUsePreview Then
                    ' Get the dimensions of the list
                    lstHwnd = GetDlgItem(GetParent(hwnd), ID_LIST)
                    GetClientRect lstHwnd, rcList
                    ClientToScreen lstHwnd, ptList
                    ClientToScreen GetParent(lstHwnd), ptWindow
                    ' Initialize the WINDOWPLACEMENT structure.
                    aWindPl.length = Len(aWindPl)
                    ' Get the dialog box's size
                    GetWindowPlacement GetParent(hwnd), aWindPl
                    ' Inflate it by about 200 pixels on the right side.
                    aWindPl.rcNormalPosition.right = aWindPl.rcNormalPosition.right + (ptList.x - ptWindow.x) + (rcList.bottom - rcList.top)
                    ' Change the size of the dialog box to our new size.
                    SetWindowPlacement GetParent(hwnd), aWindPl
                    'Create a new label
                    lPvwLbl = CreateWindowEx(WS_EX_TRANSPARENT, _
                                             "STATIC", "Preview:", WS_CHILD, _
                                             rcList.right + rcList.left + (ptList.x - ptWindow.x) * 2, _
                                             rcList.top + (ptList.y - ptWindow.y) - 22, _
                                             (aWindPl.rcNormalPosition.right - aWindPl.rcNormalPosition.left - rcList.right - (ptList.x - ptWindow.x) * 3 - 4), _
                                             20, _
                                             GetParent(hwnd), 0, App.hInstance, ByVal 0&)
                    
                    'Create the preview area
                    lPvwArea = CreateWindowEx(WS_EX_TRANSPARENT Or WS_EX_CLIENTEDGE, _
                                              "STATIC", "", WS_CHILD, _
                                              rcList.right + rcList.left + (ptList.x - ptWindow.x) * 2, _
                                              rcList.top + (ptList.y - ptWindow.y) - 2, _
                                              (aWindPl.rcNormalPosition.right - aWindPl.rcNormalPosition.left - rcList.right - (ptList.x - ptWindow.x) * 3 - 4), _
                                              (rcList.bottom - rcList.top + 4), _
                                              GetParent(hwnd), 0, App.hInstance, ByVal 0&)
                    'Create the preview dimensions box
                    lPvwInfo = CreateWindowEx(WS_EX_TRANSPARENT, _
                                             "STATIC", " ", WS_CHILD, _
                                             rcList.right + rcList.left + (ptList.x - ptWindow.x) * 2, _
                                             rcList.bottom + (ptList.y - ptWindow.y) + 4, _
                                             (aWindPl.rcNormalPosition.right - aWindPl.rcNormalPosition.left - rcList.right - (ptList.x - ptWindow.x) * 3 - 4), _
                                             (aWindPl.rcNormalPosition.bottom - rcList.bottom), _
                                             GetParent(hwnd), 0, App.hInstance, ByVal 0&)
                    'Show our label and preview area
                    ShowWindow lPvwLbl, SW_NORMAL
                    ShowWindow lPvwArea, SW_NORMAL
                    ShowWindow lPvwInfo, SW_NORMAL
                    'Store the handles against the dialog's handle
                    SetProp hwnd, "hPreviewLabel", lPvwLbl
                    SetProp hwnd, "hPreviewArea", lPvwArea
                    SetProp hwnd, "hPreviewInfo", lPvwInfo
                    SetProp lPvwArea, "hPreviewInfo", lPvwInfo
                    'Start subclassing the preview area so we can draw in the pictures.
                    SetProp lPvwArea, "OrigWndProc", SetWindowLong(lPvwArea, GWL_WNDPROC, AddressOf PreviewWndProc)
                    SetProp lPvwInfo, "OrigWndProc", SetWindowLong(lPvwInfo, GWL_WNDPROC, AddressOf PreviewInfoWndProc)
                End If
            ElseIf (anOFNotify.code = CDN_SELCHANGE) And mbUsePreview Then
                ReDim aStrAr(255)
                ' Get the selected item's path information
                strLen = SendMessage(GetParent(hwnd), CDM_GETFILEPATH, ByVal 255, aStrAr(0))
                aStr = left$(StrConv(aStrAr, vbUnicode), strLen - 1)
                If Dir(aStr, vbNormal) <> "" Or Dir(aStr, vbHidden) <> "" Then
                    If (InStr(1, aStr, ".bmp", vbTextCompare) > 0 Or _
                        InStr(1, aStr, ".gif", vbTextCompare) > 0 Or _
                        InStr(1, aStr, ".jpg", vbTextCompare) > 0) Then
                        msPvwImgPath = aStr
                        Set moPvwImg = LoadPicture(aStr)
                    Else
                        msPvwImgPath = ""
                        Set moPvwImg = Nothing
'*** by nr
                    'this working only if Preview window is set
                    'in XP this not work, Win2k is ok
                    Call SetView(GetParent(hwnd), meDesiredView)
'*** end by nr
                    End If
                Else
                    msPvwImgPath = ""
                    Set moPvwImg = Nothing
                End If
                ' Redraw both the preview area and preview info since a new item
                '  was selected
                RedrawWindow GetProp(hwnd, "hPreviewArea"), ByVal 0&, ByVal 0&, &H1
                RedrawWindow GetProp(hwnd, "hPreviewInfo"), ByVal 0&, ByVal 0&, &H1
            End If
            OFNCallbackProc = 0
        Case WM_DESTROY
            If mbUsePreview Then
                lPvwLbl = GetProp(hwnd, "hPreviewLabel")
                lPvwArea = GetProp(hwnd, "hPreviewArea")
                lPvwInfo = GetProp(hwnd, "hPreviewInfo")
                ' Destroy our label
                DestroyWindow lPvwLbl
                ' Destroy the preview area
                DestroyWindow lPvwArea
                ' destroy the preview dimensions window
                DestroyWindow lPvwInfo
                'Remove our stored values
                RemoveProp lPvwArea, "hPreviewInfo"
                RemoveProp hwnd, "hPreviewLabel"
                RemoveProp hwnd, "hPreviewArea"
                RemoveProp hwnd, "hPreviewInfo"
            End If
            Set moPvwImg = Nothing
        End Select

'*** by nr 28.01.2004   'new loop
'here is code for starting Common Dialog
'with Thumbnails Preview  (View Menu - Thumbnails)
    Dim r As RECT, p As POINTAPI
    Select Case uMsg
'      Case WM_INITDIALOG
'        'Debug.Print "WM_INITDIALOG"
'        'GetWindowRect hWndParent, r
'
'        'it can be this way
'        Call SetView(GetParent(hwnd), vsDetails)    'vsThumbnails) ' vsDetails)
      Case WM_INITDIALOG
       'by nr we will do here only thumbnails
       If (meDesiredView = vsThumbnails) Then
         'meDesiredView
        ' Initialize the WINDOWPLACEMENT structure.
        aWindPl.length = Len(aWindPl)
        ' Get the dialog box's size
        GetWindowPlacement GetParent(hwnd), aWindPl

        'hard coded values, allways the same
        p.x = aWindPl.rcNormalPosition.left + 372   '(570 - 198)
        p.y = aWindPl.rcNormalPosition.top + 37 '(145 - 108)

        SetCursorPos p.x, p.y
        mouse_event MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0
    'by nr, thumbnails is 1/5 in View Menu on Win XP
    'but 5/5 on Win2K,
            
    '         View Menu
    '   Win2000        WinXP
    '----------------------------
    '1  SmallIcons     Thumbnails
    '2  LargeIcons     Tiles
    '3  List           Icons
    '4  Details        List
    '5  Thumbnails     Details

    'xp  1/5 =  +24 'thumbnails
    'xp and 2k  3/5 = + 24 + 32  'list view
    '2k  5/5 = +24 + 72     'thumbnails
    '1/5 = +24, 2/5 = +24+16, 3/5 = +24+(16*2)
    '4/5 = +24 +(16*3),  5/5 = +24 +(16*4)
'    SetCursorPos p.x, p.y + 24            '1/5 'smallicons ?
'    SetCursorPos p.x, p.y + 24 + 16       '2/5 'largeicons ?
'    SetCursorPos p.x, p.y + 24 + (16 * 2) '3/5 'list
'    SetCursorPos p.x, p.y + 24 + (16 * 3) '4/5 'details
'    SetCursorPos p.x, p.y + 24 + (16 * 4) '5/5 'thumbnails ?
        SetCursorPos p.x, p.y + 24 + (16 * gnViewMenuPos) '5/5
        mouse_event MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0
'''        mouse_event MOUSEEVENTF_LEFTUP, p.x, p.y + 24 + (16 * gnViewMenuPos), 0, 0
        If (gnViewMenuPos = 0) Then     'xp
            mouse_event MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0
            mouse_event MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0
            'Call SetFocus(hWndParent)
            'set cursor back on View Menu
            SetCursorPos p.x, p.y
        ElseIf (gnViewMenuPos = 4) Then     'win2000
            'set cursor back on View Menu
            SetCursorPos p.x, p.y
        End If
       End If
    End Select
'*** end by nr  for thumbnails
End Function

'*********************************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************
' Subclassing procedure stuff
'*********************************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************

Private Function PreviewWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim origProc As Long    ' Our original process handle for the preview window
    
    If hwnd = 0 Then Exit Function
    
    ' Get our stored process address
    origProc = GetProp(hwnd, "OrigWndProc")
    
    Select Case message
        Case WM_PAINT
            If Not moPvwImg Is Nothing Then
                ' we are previewing and there is a valid image... draw it
                SetPreview hwnd
            Else
                ' No valid image to display... call the default process (which just
                '  erases the background to the same color as the dialog window)
                PreviewWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
            End If
        Case WM_DESTROY
            ' Unsubclass the preview window so we don't crash.
            SetWindowLong hwnd, GWL_WNDPROC, origProc
            PreviewWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
            RemoveProp hwnd, "OrigWndProc"
        Case Else
            ' Invoke the default processing. We're not interested in modifying this message.
            PreviewWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
    End Select

End Function
Private Function PreviewInfoWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim origProc    As Long         ' Original process address
    Dim aPS         As PAINTSTRUCT  ' A structure used to do custom painting in WM_PAINT
    Dim aBrush      As Long         ' A brush used to paint the background color
    Dim aBMP        As BITMAP       ' A BITMAP structure used to get the bitmap size
    Dim aStr        As String       ' Information string displayed in the preview info window
    Dim aFont       As Long         ' Original font created with the info window
    Dim aRect       As RECT         ' a RECT structure used for drawing.
    
    If hwnd = 0 Then Exit Function
    
    ' Get the original process address for the info window
    origProc = GetProp(hwnd, "OrigWndProc")
    
    Select Case message
        Case WM_PAINT
            ' Check and see if the user selected a valid image.
            If Not moPvwImg Is Nothing Then
                ' Get the dimensions of the info window
                GetClientRect hwnd, aRect
                ' Start the drawing
                BeginPaint hwnd, aPS
                ' fill in the background
                aBrush = CreateSolidBrush(GetSysColor(15))
                FillRect aPS.hdc, aRect, aBrush
                DeleteObject aBrush
                ' get the bitmap dimensions
                GetObject moPvwImg.Handle, Len(aBMP), aBMP
                ' Generate the "info" string to display
                aStr = aStr & "Pixels: " & aBMP.bmWidth & " x " & aBMP.bmHeight & vbCrLf
                aStr = aStr & "Size: " & Format(FileLen(msPvwImgPath) / 1024, "0.0 KB") & vbCrLf
                aStr = aStr & "Created: " & Format(FileDateTime(msPvwImgPath), "DD-MMM-YY HH:MM") & vbCrLf
                ' Make our text print transparent
                SetBkMode aPS.hdc, 1
                ' Draw the text with an appropriate sized font...
                aFont = SelectObject(aPS.hdc, CreateFont(-11, 0, 0, 0, 400, False, False, False, 1, 0, 0, 2, 0, "MS Sans Serif"))
                DrawText aPS.hdc, aStr, Len(aStr), aRect, DT_LEFT Or DT_TOP Or DT_WORDBREAK
                DeleteObject SelectObject(aPS.hdc, aFont)
                ' End the drawing
                EndPaint hwnd, aPS
            Else
                PreviewInfoWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
            End If
        Case WM_DESTROY
            ' Unsubclass so we don't crash out VB.
            SetWindowLong hwnd, GWL_WNDPROC, origProc
            PreviewInfoWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
            RemoveProp hwnd, "OrigWndProc"
        Case Else
            PreviewInfoWndProc = CallWindowProc(origProc, hwnd, message, wParam, lParam)
    End Select

End Function
'*********************************************************************
' Displays the clicked theme in the "preview" window
'*********************************************************************
Private Sub SetPreview(hPvwArea As Long)
        
    Dim aWid        As Long         ' Width of the preview image
    Dim aHgt        As Long         ' height of the preview image
    Dim x           As Long         ' X location in the preview window to display the image
    Dim y           As Long         ' Y location in the preview window to display the image
    Dim aBMP        As BITMAP       ' A structure used to get the image dimensions in pixels
    Dim picDC       As Long         ' A device context to hold the picture in
    Dim origBmp     As Long         ' the original 1x1 bitmap created with picDC
    Dim pvwDC       As Long         ' The preview window's device context (from the PAINTSTRUCT)
    Dim perScale    As Double       ' Percent to scale the image down, if its too big to fit
    Dim aPS         As PAINTSTRUCT  ' A paint struct used for custom processing of WM_PAINT
    Dim aBrush      As Long         ' A brush used to paint the background color of the preview area
    Dim aRect       As RECT         ' The bounding rectange for the entire preview area
    
    If hPvwArea = 0 Then Exit Sub
    
    ' Get the WHOLE area... not just the update section
    GetClientRect hPvwArea, aRect
    ' Start our custom processing of WM_PAINT
    BeginPaint hPvwArea, aPS
    ' Fill in the background with our "button face" color
    aBrush = CreateSolidBrush(GetSysColor(15))
    FillRect aPS.hdc, aRect, aBrush
    DeleteObject aBrush
    
    pvwDC = aPS.hdc
    ' Get bitmap dimensions in pixels
    GetObject moPvwImg.Handle, Len(aBMP), aBMP
    aWid = aBMP.bmWidth
    aHgt = aBMP.bmHeight
    
    ' Blt the picture
    picDC = CreateCompatibleDC(pvwDC)
    origBmp = SelectObject(picDC, moPvwImg.Handle)
    
    ' Determine scaling factor
    perScale = (aRect.right) / aWid
    If ((aRect.bottom) / aHgt) < perScale Then perScale = ((aRect.bottom) / aHgt)
    If perScale > 1 Then perScale = 1
    x = (aRect.right - CLng(aWid * perScale)) / 2
    y = (aRect.bottom - CLng(aHgt * perScale)) / 2
    
    ' Determine drawing method and transfer the image into our preview area
    If perScale < 1 Then
        SetStretchBltMode pvwDC, 4
        StretchBlt pvwDC, x, y, CLng(aWid * perScale), CLng(aHgt * perScale), picDC, 0, 0, aWid, aHgt, vbSrcCopy
    Else
        BitBlt pvwDC, x, y, aWid, aHgt, picDC, 0, 0, vbSrcCopy
    End If
       
    'Clean up our temporary DC and its 1x1 bitmap
    SelectObject picDC, origBmp
    DeleteDC picDC
    DeleteObject origBmp
    
    ' Give the DC back to its owner
    EndPaint hPvwArea, aPS

End Sub
