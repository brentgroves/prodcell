VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOwnerDrawButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     cOwnerDrawButton
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Requires: SSUBTMR.DLL
'
' Copyright � 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Draws the face of any VB button with the style set
' to graphical.  This particular implementation draws
' the standard windows buttons.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

Implements ISubclass

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_DRAWITEM = &H2B
Private Const ODT_BUTTON = 4

' Structures:
Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type

' Button messages:
Private Const BM_GETCHECK = &HF0&
Private Const BM_SETCHECK = &HF1&
Private Const BM_GETSTATE = &HF2&
Private Const BM_SETSTATE = &HF3&
Private Const BM_SETSTYLE = &HF4&
Private Const BM_CLICK = &HF5&
Private Const BM_GETIMAGE = &HF6&
Private Const BM_SETIMAGE = &HF7&

Private Const BST_UNCHECKED = &H0&
Private Const BST_CHECKED = &H1&
Private Const BST_INDETERMINATE = &H2&
Private Const BST_PUSHED = &H4&
Private Const BST_FOCUS = &H8&

' GDI etc
Private Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, ByVal eFlag As Long, ByVal eStyle As Long) As Long
Private Const DFC_CAPTION = 1
Private Const DFC_MENU = 2
Private Const DFC_SCROLL = 3
Private Const DFC_BUTTON = 4
'#if(WINVER >= =&H0500)
Private Const DFC_POPUPMENU = 5
'#endif /* WINVER >= =&H0500 */

Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONHELP = &H4

Private Const DFCS_MENUARROW = &H0
Private Const DFCS_MENUCHECK = &H1
Private Const DFCS_MENUBULLET = &H2
Private Const DFCS_MENUARROWRIGHT = &H4
Private Const DFCS_SCROLLUP = &H0
Private Const DFCS_SCROLLDOWN = &H1
Private Const DFCS_SCROLLLEFT = &H2
Private Const DFCS_SCROLLRIGHT = &H3
Private Const DFCS_SCROLLCOMBOBOX = &H5
Private Const DFCS_SCROLLSIZEGRIP = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT = &H10

Private Const DFCS_BUTTONCHECK = &H0
Private Const DFCS_BUTTONRADIOIMAGE = &H1
Private Const DFCS_BUTTONRADIOMASK = &H2
Private Const DFCS_BUTTONRADIO = &H4
Private Const DFCS_BUTTON3STATE = &H8
Private Const DFCS_BUTTONPUSH = &H10

Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const DFCS_CHECKED = &H400

'#if(WINVER >= =&H0500)
Private Const DFCS_TRANSPARENT = &H800
Private Const DFCS_HOT = &H1000
'#endif /* WINVER >= =&H0500 */

Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_FLAT = &H4000
Private Const DFCS_MONO = &H8000

Public Enum EODStandardButtonStyle
   eodNone = 0
   eodClose = 1
   eodHelp = 2
   eodMax = 3
   eodMin = 4
   eodRestore = 5
   eodDown = 6
   eodUp = 7
   eodLeft = 8
   eodRight = 9
   eodSizeBox = 10
End Enum

Private Const vbalODButtonStyleProp = "vbalODButtonStyle"

Private m_hWnd As Long

Public Sub Attach(ByRef hWndButtonOwner As Long)
   Detach
   m_hWnd = hWndButtonOwner
   AttachMessage Me, m_hWnd, WM_DRAWITEM
End Sub
Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_DRAWITEM
      m_hWnd = 0
   End If
End Sub
Public Property Let ButtonStyle(ByVal hwnd As Long, ByVal eStyle As EODStandardButtonStyle)
   If eStyle < eodNone Then
      RemoveProp hwnd, vbalODButtonStyleProp
   Else
      SetProp hwnd, vbalODButtonStyleProp, eStyle
   End If
End Property
Public Property Get ButtonStyle(ByVal hwnd As Long) As EODStandardButtonStyle
   ButtonStyle = GetProp(hwnd, vbalODButtonStyleProp)
End Property

Private Sub pDrawButton(tDis As DRAWITEMSTRUCT, eStyle As EODStandardButtonStyle)
Dim hBr As Long
Dim lState As Long
Dim bPushed As Boolean
Dim bEnabled As Boolean
Dim bChecked As Boolean
Dim bFocus As Boolean
Dim lType As Long
Dim lStyle As Long
   
   ' Find out about the button:
   lState = SendMessageLong(tDis.hwndItem, BM_GETSTATE, 0, 0)
   bPushed = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
   bChecked = (SendMessageLong(tDis.hwndItem, BM_GETCHECK, 0, 0) <> 0)
   bEnabled = IsWindowEnabled(tDis.hwndItem)
   bFocus = (GetFocus() = tDis.hwndItem)

   ' Use draw frame control to render the button:
   Select Case eStyle
   Case eodUp
      lType = DFC_SCROLL
      lStyle = DFCS_SCROLLUP
   Case eodSizeBox
      lType = DFC_SCROLL
      lStyle = DFCS_SCROLLSIZEGRIP
   Case eodRight
      lType = DFC_SCROLL
      lStyle = DFCS_SCROLLRIGHT
   Case eodRestore
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONRESTORE
   Case eodMin
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONMIN
   Case eodMax
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONMAX
   Case eodLeft
      lType = DFC_SCROLL
      lStyle = DFCS_SCROLLLEFT
   Case eodHelp
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONHELP
   Case eodDown
      lType = DFC_SCROLL
      lStyle = DFCS_SCROLLDOWN
   Case eodClose
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONCLOSE
   End Select
   If (bPushed) Or (bChecked) Then
      lStyle = lStyle Or DFCS_PUSHED
   End If
   If Not (bEnabled) Then
      lStyle = lStyle Or DFCS_INACTIVE
   End If
   DrawFrameControl tDis.hdc, tDis.rcItem, lType, lStyle


End Sub

Private Sub Class_Terminate()
   Detach
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If (CurrentMessage = WM_DRAWITEM) Then
      ISubclass_MsgResponse = emrConsume
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDis As DRAWITEMSTRUCT
Dim eStyle As EODStandardButtonStyle

   Select Case iMsg
   Case WM_DRAWITEM
      CopyMemory tDis, ByVal lParam, Len(tDis)
      If tDis.CtlType = ODT_BUTTON Then
         eStyle = GetProp(tDis.hwndItem, vbalODButtonStyleProp)
         If (eStyle > 0) Then
            pDrawButton tDis, eStyle
            ISubclass_WindowProc = 1
         Else
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If
      End If
   End Select
   
End Function


