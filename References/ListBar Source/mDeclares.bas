Attribute VB_Name = "mDeclares"
Option Explicit

' ======================================================================================
' Name:     mDeclares.Bas
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Requires: None
'
' Copyright � 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' This file contains the declares for the
' vbAccelerator ListBar control.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type NMHDR
   hwndFrom As Long
   idfrom As Long
   code As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

Public Const SW_HIDE = 0

' Some styles:
Public Const GWL_STYLE = (-16)
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_DISABLED = &H8000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_TABSTOP = &H100000
Public Const WS_HSCROLL = &H100000
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000
Public Const CW_USEDEFAULT = &H80000000

Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const WM_SETFOCUS = &H7
Public Const WM_ERASEBKGND = &H14
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_NOTIFY = &H4E
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_PARENTNOTIFY = &H210

' mouse activate responses
Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4

Public Const CLR_INVALID = -1
Public Const CLR_NONE = &HFFFFFFFF

' GDI object functions:
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const BITSPIXEL = 12
    Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
' System metrics:
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Const SM_CXICON = 11
    Public Const SM_CYICON = 12
    Public Const SM_CXFRAME = 32
    Public Const SM_CYCAPTION = 4
    Public Const SM_CYFRAME = 33
    Public Const SM_CYHSCROLL = 3
    Public Const SM_CYBORDER = 6
    Public Const SM_CXBORDER = 5
    Public Const SM_CYMENU = 15
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
    
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
             Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE = 2
    Public Const TRANSPARENT = 1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Public Const DT_BOTTOM = &H8&
    Public Const DT_CENTER = &H1&
    Public Const DT_LEFT = &H0&
    Public Const DT_CALCRECT = &H400&
    Public Const DT_WORDBREAK = &H10&
    Public Const DT_VCENTER = &H4&
    Public Const DT_TOP = &H0&
    Public Const DT_TABSTOP = &H80&
    Public Const DT_SINGLELINE = &H20&
    Public Const DT_RIGHT = &H2&
    Public Const DT_NOCLIP = &H100&
    Public Const DT_INTERNAL = &H1000&
    Public Const DT_EXTERNALLEADING = &H200&
    Public Const DT_EXPANDTABS = &H40&
    Public Const DT_CHARSTREAM = 4&
    Public Const DT_WORD_ELLIPSIS = &H40000
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Const PS_SOLID = 0

' Cursor functions:
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const IDC_ARROW = 32512&
Public Const WM_SETCURSOR = &H20
Public Const GCL_HCURSOR = (-12)

' Rectangle functions:
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

' Get version:
Public Type DLLVERSIONINFO
   cbSize As Long
   dwMajor As Long
   dwMinor As Long
   dwBuildNumber As Long
   dwPlatformID As Long
End Type
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long

' ImageList
Public Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cY As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_DrawEx Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Public Const ILD_TRANSPARENT = 1&



