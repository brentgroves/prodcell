VERSION 5.00
Begin VB.UserControl vbalListBar 
   Alignable       =   -1  'True
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   MouseIcon       =   "vbalListBar.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5535
   ScaleWidth      =   2880
   ToolboxBitmap   =   "vbalListBar.ctx":0152
   Begin VB.Timer tmrScrolling 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   1380
      Top             =   4740
   End
   Begin VB.PictureBox picScroll 
      Height          =   3795
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   2235
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   660
      Width           =   2295
      Begin VB.CommandButton cmdDown 
         Height          =   555
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdUp 
         Height          =   555
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox picLvw 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   60
         ScaleHeight     =   3615
         ScaleWidth      =   2115
         TabIndex        =   3
         Top             =   60
         Width           =   2115
      End
   End
End
Attribute VB_Name = "vbalListBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ======================================================================================
' Name:     vbalListBar.ctl
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Requires: SSUBTMR.DLL,
'
'
' Copyright � 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' ListBar control implementation.
'
' Changes:
' SPM 2 February 1999
' * Fixed item alignment problem for large icons mode (items did not
'   always centre)
'   Thanks to Rafael <rapm20@cantv.net>
' * Fixed TrackSelect problem - code was in reverse...
'   Thanks to Andr�s Giraldo <andres_giraldo@yahoo.com>
' * Fixed cursors - now the hand shows for the ListBar buttons and
'   the arrow cursor shows for the items within a bar rather than
'   vice-versa
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

' Enums:
Public Enum EVBALLBBSortOrderConstants
   [_First] = 0
   evlbAscending = 0
   evlbDescending = 1
   [_Last] = 1
End Enum

Public Enum EVBALLBIconSizeConstants
   evlbLargeIcon = 0
   evlbSmallIcon = 1
End Enum

Public Enum EVBALLBBBorderStyleConstants
   [_First] = 0
   evlbNone = 0
   evlb3D = 1
   evlb3DThin = 2
   [_Last] = 2
End Enum



Private m_hWndCtl As Long
Private m_bRunTime As Boolean

Private Type tListBarItem
   iBarID As Long
   iItemID As Long
   sCaption As String
   sHelpText As String
   sKey As String
   lItemData As Long
   lIconIndex As Long
   sTag As String
   bInUse As Boolean
End Type

Private Type tListBar
   iID As Long
   sCaption As String
   sHelpText As String
   sTag As String
   lItemData As Long
   sToolTip As String
   sKey As String
   eIconSize As EVBALLBIconSizeConstants
   eView As Long 'EVBALLViewStyleConstants
   oBackColor As OLE_COLOR
   oForeColor As OLE_COLOR
   oHighlightColor As OLE_COLOR
   bOfficeXpStyle As Boolean
   tR As RECT
   bHot As Boolean
   bPressed As Boolean
   lTop As Long
   tItems() As tListBarItem
   lItemCount As Long
   lItemSelected As Long
End Type

Private m_tBars() As tListBar
Private m_iBarCount As Long
Private m_iSelBar As Long
Private m_iButtonHeight As Long
Private m_lItemHeight As Long
Private m_iMouseDownBtn As Long
Private m_iDownOn As Long
Private m_iRDownOn As Long

Private m_lEditBar As Long
Private m_lEditItem As Long
Private m_bDragging As Boolean
Private m_tDragBegin As POINTAPI

Private m_cMemDC As cMemDC
Private m_cAnimDC As cMemDC

Private m_oForeColor As OLE_COLOR
Private m_oBackColor As OLE_COLOR
Private m_sBackgroundPicture As String

Private m_fnt As IFont
Private WithEvents m_cTPM As cMouseTrack
Attribute m_cTPM.VB_VarHelpID = -1
Private m_cODBtn As cOwnerDrawButton

Private m_hIml(0 To 1) As Long
Private m_lIconSizeX(0 To 1) As Long
Private m_lIconSizeY(0 To 1) As Long

Private m_eBorderStyle As EVBALLBBBorderStyleConstants

Private m_bDrag As Boolean

' Over-riding VB UserControl's default IOLEInPlaceActivate:
Private m_IPAOHookStruct As IPAOHookStruct

Private WithEvents lvw As cListView
Attribute lvw.VB_VarHelpID = -1
Private m_hWndLV As Long

Implements ISubclass

'Public Event ItemClick(ByVal lBarIndex As Long, ByVal lItemIndex As Long)
'Public Event BarClick(ByVal lBar As Long)

Public Event BarClick(Bar As cListBar)
Attribute BarClick.VB_Description = "Raised when a bar is selected."
Public Event ItemClick(Item As cListBarItem, Bar As cListBar)
Attribute ItemClick.VB_Description = "Raised when a Item within the ListBar is clicked."
Public Event ItemEndEdit(Item As cListBarItem, ByRef sText As String, ByRef bCancel As Boolean)
Attribute ItemEndEdit.VB_Description = "Raised when the user completes editing an item in the bar."
Public Event ItemRightClick(Item As cListBarItem, Bar As cListBar, x As Single, y As Single)
Attribute ItemRightClick.VB_Description = "Raised when the user right clicks an item in the bar."
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Raised when the user depresses the mouse button on a non-active part of the control (i.e. not a bar selector or an item)"
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Raised when the user releases the mouse button on a non-active part of the control (i.e. not a bar selector or an item)"

Friend Function TranslateAccelerator(lpMsg As VBOleGuids.MSG) As Long
   If Not lvw Is Nothing Then
      TranslateAccelerator = lvw.TranslateAccelerator(lpMsg)
   End If
End Function

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gets the Window Handle of the control."
   hWnd = m_hWndCtl
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the foreground color of the control ( bar selection button text)"
   ForeColor = m_oForeColor
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
   m_oForeColor = oColor
   UserControl.ForeColor = oColor
   Render
   PropertyChanged "ForeColor"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the background color of the control (appears behind the bar selection buttons)"
   BackColor = m_oBackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
   m_oBackColor = oColor
   UserControl.BackColor = oColor
   Render
   PropertyChanged "BackColor"
End Property

Public Property Get Font() As IFont
Attribute Font.VB_Description = "Gets/sets the font to draw the control."
   Set Font = UserControl.Font
End Property

Public Property Let Font(fnt As IFont)
   pSetFont fnt
   PropertyChanged "Font"
End Property

Public Property Set Font(fnt As IFont)
   pSetFont fnt
   PropertyChanged "Font"
End Property

Private Sub pSetFont(fnt As IFont)
   Set UserControl.Font = fnt
   Set m_fnt = fnt
   Render
End Sub

Public Property Get BackgroundPicture() As String
Attribute BackgroundPicture.VB_Description = "Gets/sets a picture to display behind the items within a bar.  Only available if you have COMCTL32.DLL v4.72 or higher."
   BackgroundPicture = m_sBackgroundPicture
End Property
Public Property Let BackgroundPicture(ByVal sURL As String)
   m_sBackgroundPicture = sURL
   lvw.BackgroundPicture = sURL
   PropertyChanged "BackgroundPicture"
End Property

Public Property Get Bars() As cListBars
Attribute Bars.VB_Description = "Returns a reference to the control's Bars collection."
   Dim cL As cListBars
   Set cL = New cListBars
   cL.fInit m_hWndCtl
   Set Bars = cL
End Property
Public Property Get SelectedBar() As cListBar
Attribute SelectedBar.VB_Description = "Gets a reference to the currently selected bar in the control."
   If m_iSelBar > 0 Then
      Dim cLB As cListBar
      Set cLB = New cListBar
      cLB.fInit m_hWndCtl, m_tBars(m_iSelBar).iID
      Set SelectedBar = cLB
   End If
End Property

Public Property Get ScaleMode() As ScaleModeConstants
   ScaleMode = UserControl.ScaleMode
End Property
Public Property Let ScaleMode(ByVal eMode As ScaleModeConstants)
   If Not (UserControl.ScaleMode = eMode) Then
      UserControl.ScaleMode = eMode
      PropertyChanged "ScaleMode"
   End If
End Property
Public Function ScaleX(x As Single, fromScale As ScaleModeConstants, toScale As ScaleModeConstants)
   ScaleX = UserControl.ScaleX(x, fromScale, toScale)
End Function
Public Function ScaleY(y As Single, fromScale As ScaleModeConstants, toScale As ScaleModeConstants)
   ScaleY = UserControl.ScaleY(y, fromScale, toScale)
End Function

Public Property Let ImageList(Optional ByVal eSize As EVBALLBIconSizeConstants = evlbLargeIcon, vThis As Variant)
Attribute ImageList.VB_Description = "Allows an ImageList to be associated with the control.  Either pass in a Microsoft ImageList control or a long value containing the hImageList handle to the ImageList."
   pImageList eSize, vThis
End Property
Public Property Set ImageList(Optional ByVal eSize As EVBALLBIconSizeConstants = evlbLargeIcon, vThis As Variant)
   pImageList eSize, vThis
End Property
Private Sub pImageList(ByVal eSize As EVBALLBIconSizeConstants, vThis As Variant)
Dim hIml As Long
   If eSize <> evlbLargeIcon And eSize <> evlbSmallIcon Then
      eSize = evlbLargeIcon
   End If

   ' Set the ImageList handle property either from a VB
   ' image list or directly:
   If VarType(vThis) = vbObject Then
       ' Assume VB ImageList control.  Note that unless
       ' some call has been made to an object within a
       ' VB ImageList the image list itself is not
       ' created.  Therefore hImageList returns error. So
       ' ensure that the ImageList has been initialised by
       ' drawing into nowhere:
       On Error Resume Next
       ' Get the image list initialised..
       vThis.ListImages(1).Draw 0, 0, 0, 1
       hIml = vThis.hImageList
       If (Err.Number <> 0) Then
           hIml = 0
       End If
       On Error GoTo 0
   ElseIf VarType(vThis) = vbLong Then
       ' Assume ImageList handle:
       hIml = vThis
   Else
       Err.Raise vbObjectError + 1049, "cToolbar." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
   End If
    
   ' If we have a valid image list, then associate it with the control:
   If (hIml <> 0) Then
      m_hIml(eSize) = hIml
      ImageList_GetIconSize m_hIml(eSize), m_lIconSizeX(eSize), m_lIconSizeY(eSize)
      If pbEnsureListView() Then
         lvw.ImageList(eSize) = vThis
      End If
   End If

End Sub

Private Function pbIsValidKey(ByVal sKey As String) As Boolean
Dim i As Long
   If IsNumeric(sKey) Then
      ' Invalid Key, numeric
      gErr 4, "vbalListBar"
   Else
      pbIsValidKey = True
      For i = 1 To m_iBarCount
         If m_tBars(i).sKey = sKey Then
            ' duplicate key
            gErr 5, "vbalListBar"
            pbIsValidKey = False
            Exit For
         End If
      Next i
   End If
End Function
Private Function pbIsValidItemKey(ByVal lBar As Long, ByVal sKey As String) As Boolean
Dim i As Long
   If IsNumeric(sKey) Then
      ' Invalid Key, numeric
      gErr 4, "vbalListBar"
   Else
      pbIsValidItemKey = True
      With m_tBars(lBar)
         For i = 1 To .lItemCount
            If .tItems(i).sKey = sKey Then
               ' duplicate key
               gErr 5, "vbalListBar"
               pbIsValidItemKey = False
               Exit For
            End If
         Next i
      End With
   End If
End Function
Friend Property Get fItemSelectedID(ByVal lBarPosition As Long) As Long
   With m_tBars(lBarPosition)
      If .lItemSelected > 0 Then
         fItemSelectedID = (.tItems(.lItemSelected).iItemID)
      End If
   End With
End Property
Friend Property Get fItemID(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As Long
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemID = .iItemID
   End With
End Property
Friend Property Get fItemTag(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As String
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemTag = .sTag
   End With
End Property
Friend Property Let fItemTag(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal sTag As String)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .sTag = sTag
   End With
End Property
Friend Property Get fItemKey(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As String
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemKey = .sKey
   End With
End Property
Friend Property Let fItemKey(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal sKey As String)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .sKey = sKey
   End With
End Property
Friend Sub fItemGetRect(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByRef tR As RECT)
   If lBarPosition = m_iSelBar Then
      lvw.GetItemRect lItemPosition, tR.Left, tR.Top, tR.Right, tR.Bottom
   Else
      ' synthesize (todo)
      If m_tBars(lBarPosition).eIconSize = evlbLargeIcon Then
         
      Else
         
      End If
   End If
End Sub
Friend Sub fBeginEdit(ByVal lBarPosition As Long, ByVal lItemPosition As Long)
   If lBarPosition = m_iSelBar Then
      lvw.EnsureVisible lItemPosition
      lvw.StartEdit lItemPosition
      m_lEditBar = lBarPosition
      m_lEditItem = lItemPosition
   End If
End Sub
Friend Property Get fItemCaption(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As String
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemCaption = .sCaption
   End With
End Property
Friend Property Let fItemCaption(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal sCaption As String)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .sCaption = sCaption
      If m_iSelBar = lBarPosition Then
         lvw.ItemCaption(lItemPosition) = sCaption
      End If
   End With
End Property
Friend Property Get fItemHelpText(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As String
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemHelpText = .sHelpText
   End With
End Property
Friend Property Let fItemHelpText(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal sHelpText As String)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .sHelpText = sHelpText
   End With
End Property
Friend Property Get fItemIconIndex(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As Long
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemIconIndex = .lIconIndex
   End With
End Property
Friend Property Let fItemIconIndex(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal lIconIndex As Long)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .lIconIndex = lIconIndex
      If m_iSelBar = lBarPosition Then
         lvw.ItemIconIndex(lItemPosition) = lIconIndex
      End If
   End With
End Property
Friend Property Get fItemItemData(ByVal lBarPosition As Long, ByVal lItemPosition As Long) As Long
   With m_tBars(lBarPosition).tItems(lItemPosition)
      fItemItemData = .lItemData
   End With
End Property
Friend Property Let fItemItemData(ByVal lBarPosition As Long, ByVal lItemPosition As Long, ByVal lItemData As Long)
   With m_tBars(lBarPosition).tItems(lItemPosition)
      .lItemData = lItemData
   End With
End Property
Friend Property Get fIsItem(ByVal lPosition As Long, ByVal lItemID As Long) As Long
Dim i As Long
   With m_tBars(lPosition)
      For i = 1 To .lItemCount
         If .tItems(i).iItemID = lItemID Then
            fIsItem = i
            Exit Property
         End If
      Next i
   End With
End Property
Friend Property Get fIsBar(ByVal lID As Long) As Long
Dim i As Long
   For i = 1 To m_iBarCount
      If m_tBars(i).iID = lID Then
         fIsBar = i
         Exit Property
      End If
   Next i
End Property
Friend Function fPositionForKey(vKey As Variant) As Long
Dim i As Long
   If IsNumeric(vKey) Then
      If vKey > 0 And vKey <= m_iBarCount Then
         fPositionForKey = vKey
      Else
         gErr 6, "vbalListBar"
      End If
   Else
      For i = 1 To m_iBarCount
         If m_tBars(i).sKey = vKey Then
            fPositionForKey = i
            Exit Function
         End If
      Next i
      gErr 6, "vbalListBar"
   End If
End Function
Friend Function fAddItem(lBar As Long, vKey As Variant, vKeyBefore As Variant, sCaption As String, lIconIndex As Long) As Long
Dim lInsIndex As Long
Dim i As Long
Dim bGenKey As Boolean
   
   If Not (IsMissing(vKeyBefore)) Then
      ' Check whether we can do that
      lInsIndex = fItemIndex(lBar, vKeyBefore)
      If lInsIndex = 0 Then
         Exit Function
      End If
   End If
   
   ' Verify key validity:
   If IsMissing(vKey) Then
      bGenKey = True
   Else
      If Not pbIsValidItemKey(lBar, vKey) Then
         Exit Function
      End If
   End If
   
   ' Everything is ok, let's add the item:
   With m_tBars(lBar)
      .lItemCount = .lItemCount + 1
      ReDim Preserve .tItems(1 To .lItemCount) As tListBarItem
      If lInsIndex > 0 Then
         For i = .lItemCount - 1 To lInsIndex Step -1
            LSet .tItems(i + 1) = .tItems(i)
         Next i
         If .lItemSelected >= lInsIndex Then
            .lItemSelected = .lItemSelected + 1
         End If
      Else
         lInsIndex = .lItemCount
      End If
      
      With .tItems(lInsIndex)
         .iBarID = m_tBars(lBar).iID
         .iItemID = gNewItemID
         .sCaption = sCaption
         .lIconIndex = lIconIndex
         If bGenKey Then
            .sKey = "C" & .iItemID
         Else
            .sKey = vKey
         End If
         fAddItem = .iItemID
      End With
      
      If lBar = m_iSelBar Then
         pSetUpBar lBar
         ' SPM 25/02/00 Bug fix
         picScroll_Resize
      End If
      
   End With

End Function
Friend Sub fClearBar(ByVal lBarIndex As Long)
Dim lS As Long
   '
   m_tBars(lBarIndex).lItemCount = 0
   Erase m_tBars(lBarIndex).tItems
   If (lBarIndex = m_iSelBar) Then
      lvw.Clear
      lvw.Update
      lS = GetWindowLong(lvw.hWndLV, GWL_STYLE)
      lS = lS And Not WS_HSCROLL
      SetWindowLong lvw.hWndLV, GWL_STYLE, lS
      picScroll_Resize
   End If
   '
End Sub
Friend Function fAddBar(vKey As Variant, vKeyBefore As Variant, sCaption As String, sToolTip As String) As Long
Dim lInsIndex As Long
Dim i As Long
Dim bGenKey As Boolean

   If Not IsNull(vKeyBefore) Then
      If Not (IsMissing(vKeyBefore)) Then
         ' Check whether we can do that
         lInsIndex = fPositionForKey(vKeyBefore)
         If lInsIndex = 0 Then
            Exit Function
         End If
      End If
   End If
   
   ' Verify key validity:
   If IsMissing(vKey) Then
      bGenKey = True
   Else
      If Not pbIsValidKey(vKey) Then
         Exit Function
      End If
   End If
   
   ' If we don't have a listview yet, time to
   ' create one:
   If pbEnsureListView() Then
   
      ' Everything is ok, let's add the item:
      m_iBarCount = m_iBarCount + 1
      ReDim Preserve m_tBars(1 To m_iBarCount) As tListBar
      If lInsIndex > 0 Then
         For i = m_iBarCount - 1 To lInsIndex Step -1
            LSet m_tBars(i + 1) = m_tBars(i)
         Next i
      Else
         lInsIndex = m_iBarCount
      End If
      
      With m_tBars(lInsIndex)
         .iID = gNewBarID
         .sCaption = sCaption
         If bGenKey Then
            .sKey = "C" & .iID
         Else
            .sKey = vKey
         End If
         .oBackColor = vbButtonShadow
         .oForeColor = vb3DHighlight
         .oHighlightColor = vbHighlight
         .sToolTip = sToolTip
      End With
      If m_iSelBar = 0 Then
         fBarSelect 1
      Else
         Render
      End If
      fAddBar = m_tBars(lInsIndex).iID
   End If
   
End Function
Friend Function fClear()
   m_iBarCount = 0
   Erase m_tBars
   m_iSelBar = 0
   Render
End Function
Friend Function fRemoveBar(vKey As Variant) As Boolean
Dim i As Long
Dim lIdx As Long
   lIdx = fPositionForKey(vKey)
   If lIdx > 0 Then
      m_iBarCount = m_iBarCount - 1
      If m_iBarCount <= 0 Then
         m_iSelBar = 0
         m_iBarCount = 0
         Erase m_tBars
      Else
         For i = lIdx To m_iBarCount
            LSet m_tBars(i) = m_tBars(i + 1)
         Next i
         ReDim Preserve m_tBars(1 To m_iBarCount) As tListBar
         If m_iSelBar > m_iBarCount Then
            fBarSelect m_iBarCount
         End If
      End If
      Render
   End If
End Function
Friend Function fRemoveItem(lBar As Long, vKey As Variant) As Boolean
Dim i As Long
Dim lIdx As Long
   lIdx = fItemIndex(lBar, vKey)
   If lIdx > 0 Then
      With m_tBars(lBar)
         .lItemCount = .lItemCount - 1
         If .lItemCount <= 0 Then
            .lItemCount = 0
            Erase .tItems
            If m_iSelBar = lBar Then
               pSetUpBar lBar
               picScroll_Resize
            End If
         Else
            For i = lIdx To .lItemCount
               LSet .tItems(i) = .tItems(i + 1)
            Next i
            ReDim Preserve .tItems(1 To .lItemCount) As tListBarItem
            If .lItemSelected > lIdx Then
               .lItemSelected = .lItemSelected - 1
            End If
            If m_iSelBar = lBar Then
               pSetUpBar lBar
               picScroll_Resize
            End If
         End If
      End With
   End If
End Function
Friend Function fBarIDForKey(vKey As Variant) As Long
Dim i As Long
   If IsNumeric(vKey) Then
      If vKey > 0 And vKey <= m_iBarCount Then
         fBarIDForKey = m_tBars(vKey).iID
      Else
         gErr 6, "vbalListBar"
      End If
   Else
      For i = 1 To m_iBarCount
         If m_tBars(i).sKey = vKey Then
            fBarIDForKey = m_tBars(i).iID
            Exit Function
         End If
      Next i
      gErr 6, "vbalListBar"
   End If
End Function
Friend Property Get fBarCount() As Long
   fBarCount = m_iBarCount
End Property
Friend Property Get fBarSelected(ByVal lPosition As Long) As Boolean
   fBarSelected = (lPosition = m_iSelBar)
End Property
Friend Property Get fBarSelectedItemID(ByVal lPosition As Long) As Long
   fBarSelectedItemID = m_tBars(lPosition).tItems(m_tBars(lPosition).lItemSelected).iItemID
End Property
Friend Property Get fBarCaption(ByVal lPosition As Long) As String
   fBarCaption = m_tBars(lPosition).sCaption
End Property
Friend Property Let fBarCaption(ByVal lPosition As Long, ByVal sCaption As String)
   m_tBars(lPosition).sCaption = sCaption
   Render
End Property
Friend Property Get fBarKey(ByVal lPosition As Long) As String
   fBarKey = m_tBars(lPosition).sKey
End Property
Friend Property Let fBarKey(ByVal lPosition As Long, ByVal sKey As String)
   If pbIsValidKey(sKey) Then
      m_tBars(lPosition).sKey = sKey
   End If
End Property
Friend Property Get fBarTag(ByVal lPosition As Long) As String
   fBarTag = m_tBars(lPosition).sTag
End Property
Friend Property Let fBarTag(ByVal lPosition As Long, ByVal sTag As String)
   m_tBars(lPosition).sTag = sTag
End Property
Friend Property Get fBarHelpText(ByVal lPosition As Long) As String
   fBarHelpText = m_tBars(lPosition).sHelpText
End Property
Friend Property Let fBarHelpText(ByVal lPosition As Long, ByVal sHelpText As String)
   m_tBars(lPosition).sHelpText = sHelpText
End Property
Friend Property Get fBarItemData(ByVal lPosition As Long) As Long
   fBarItemData = m_tBars(lPosition).lItemData
End Property
Friend Property Let fBarItemData(ByVal lPosition As Long, ByVal lItemData As Long)
   m_tBars(lPosition).lItemData = lItemData
End Property
Friend Property Get fBarBackColor(ByVal lPosition As Long) As OLE_COLOR
   fBarBackColor = m_tBars(lPosition).oBackColor
End Property
Friend Property Let fBarBackColor(ByVal lPosition As Long, ByVal oColor As OLE_COLOR)
   m_tBars(lPosition).oBackColor = oColor
   If m_iSelBar = lPosition Then
      lvw.BackColor = oColor
      Dim tR As RECT
      GetClientRect picLvw.hWnd, tR
      RedrawWindow picLvw.hWnd, tR, 0, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_INVALIDATE Or RDW_UPDATENOW
   End If
End Property
Friend Property Get fBarForeColor(ByVal lPosition As Long) As OLE_COLOR
   fBarForeColor = m_tBars(lPosition).oForeColor
End Property
Friend Property Let fBarForeColor(ByVal lPosition As Long, ByVal oColor As OLE_COLOR)
   m_tBars(lPosition).oForeColor = oColor
   If m_iSelBar = lPosition Then
      lvw.ForeColor = oColor
      Dim tR As RECT
      GetClientRect picLvw.hWnd, tR
      RedrawWindow picLvw.hWnd, tR, 0, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_INVALIDATE Or RDW_UPDATENOW
   End If
End Property
Friend Property Get fBarHighlightColor(ByVal lPosition As Long) As OLE_COLOR
   fBarHighlightColor = m_tBars(lPosition).oHighlightColor
End Property
Friend Property Let fBarHighlightColor(ByVal lPosition As Long, ByVal oColor As OLE_COLOR)
   m_tBars(lPosition).oHighlightColor = oColor
   If m_iSelBar = lPosition Then
      lvw.HighlightColor = oColor
   End If
End Property
Friend Property Get fBarOfficeXpStyle(ByVal lPosition As Long) As Boolean
   fBarOfficeXpStyle = m_tBars(lPosition).bOfficeXpStyle
End Property
Friend Property Let fBarOfficeXpStyle(ByVal lPosition As Long, ByVal bState As Boolean)
   m_tBars(lPosition).bOfficeXpStyle = bState
   If m_iSelBar = lPosition Then
      lvw.OfficeXpStyle = bState
   End If
End Property

Friend Property Get fBarIconSize(ByVal lPosition As Long) As EVBALLBIconSizeConstants
   fBarIconSize = m_tBars(lPosition).eIconSize
End Property
Friend Property Let fBarIconSize(ByVal lPosition As Long, ByVal eSize As EVBALLBIconSizeConstants)
   m_tBars(lPosition).eIconSize = eSize
   pSetUpBar lPosition
   If m_bRunTime Then
      picScroll_Resize
   End If
End Property
Friend Sub fBarSelect(ByVal lPosition As Long)
   ' Select the bar
   m_iSelBar = lPosition
   picLvw.Top = m_tBars(lPosition).lTop
   pSetUpBar lPosition
   Render
End Sub
Friend Sub fBarItemSelect(ByVal lBarPosition As Long, ByVal lItemPosition As Long)
   pSelectItem lBarPosition, lItemPosition
End Sub
Friend Sub fBarSort(ByVal lPosition As Long, ByVal eDir As EVBALLBBSortOrderConstants)
   ' TODO!
   
   ' Re-render the bar:
   pSetUpBar lPosition
   Render
End Sub
Friend Property Get fItemCount(ByVal lPosition As Long) As Long
   fItemCount = m_tBars(lPosition).lItemCount
End Property
Friend Property Get fItemIndex(ByVal lPosition As Long, vKey As Variant) As Long
Dim i As Long
   With m_tBars(lPosition)
      If IsNumeric(vKey) Then
         If vKey > 0 And vKey <= .lItemCount Then
            fItemIndex = vKey
         Else
            gErr 7, "vbalListBar"
         End If
      Else
         For i = 1 To .lItemCount
            If .tItems(i).sKey = vKey Then
               fItemIndex = i
               Exit Property
            End If
         Next i
         gErr 7, "vbalListBar"
      End If
   End With
End Property
Public Property Get BorderStyle() As EVBALLBBBorderStyleConstants
Attribute BorderStyle.VB_Description = "Gets/sets the type of border to show around the control."
   BorderStyle = m_eBorderStyle
End Property
Public Property Let BorderStyle(ByVal eStyle As EVBALLBBBorderStyleConstants)
Dim lhWnd As Long
Dim lS As Long
   m_eBorderStyle = eStyle
   If eStyle = evlbNone Then
      UserControl.BorderStyle() = 0
   Else
      UserControl.BorderStyle() = (eStyle = evlb3D)
      lhWnd = UserControl.hWnd
         lS = GetWindowLong(lhWnd, GWL_EXSTYLE)
      If eStyle = evlb3D Then
         lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
      Else
         lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
      End If
      SetWindowLong lhWnd, GWL_EXSTYLE, lS
      SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
   End If
   PropertyChanged "BorderStyle"
End Property

Private Sub pDrawButton(ByVal lhDCTo As Long, ByVal bMemDC As Boolean, ByVal iButton As Long)
Dim tTR As RECT
Dim tR As RECT
Dim tTROut As RECT
Dim dwFlags As Long
Dim hBr As Long
Dim hFontOld As Long

   GetClientRect m_hWndCtl, tTR
   tTR.Top = m_iSelBar * m_iButtonHeight
   tTR.Bottom = tTR.Bottom - (m_iBarCount - m_iSelBar) * m_iButtonHeight
   GetWindowRect picScroll.hWnd, tR
   If (EqualRect(tTR, tR) = 0) Then
      'SetWindowPos lvw.hwnd, 0, tTR.left, tTR.top, tTR.right - tTR.left, tTR.bottom - tTR.top, SWP_NOZORDER Or SWP_NOOWNERZORDER
      If tTR.Bottom - tTR.Top > 0 And tTR.Right - tTR.Left > 0 Then
         picScroll.Move tTR.Left * Screen.TwipsPerPixelX, tTR.Top * Screen.TwipsPerPixelY, (tTR.Right - tTR.Left) * Screen.TwipsPerPixelX, (tTR.Bottom - tTR.Top) * Screen.TwipsPerPixelY
      Else
         ' ...
      End If
   End If

   LSet m_tBars(iButton).tR = tTR
   If iButton <= m_iSelBar Then
      ' drawing from top:
      m_tBars(iButton).tR.Top = (iButton - 1) * m_iButtonHeight
   Else
      ' drawing from bottom:
      m_tBars(iButton).tR.Top = tTR.Bottom + (iButton - m_iSelBar - 1) * m_iButtonHeight
   End If
   m_tBars(iButton).tR.Bottom = m_tBars(iButton).tR.Top + m_iButtonHeight
   
   LSet tTROut = m_tBars(iButton).tR
   If m_tBars(iButton).bPressed Then
      If m_tBars(iButton).bHot Then
         dwFlags = EDGE_SUNKEN
      Else
         dwFlags = BDR_SUNKENINNER
      End If
   Else
      If m_tBars(iButton).bHot Then
         dwFlags = EDGE_RAISED
      Else
         dwFlags = BDR_RAISEDINNER
      End If
   End If
   dwFlags = dwFlags Or BF_SOFT Or BF_MIDDLE
   If bMemDC Then
      OffsetRect tTROut, 0, -tTROut.Top
   End If
   hBr = CreateSolidBrush(TranslateColor(m_oBackColor))
   FillRect lhDCTo, tTROut, hBr
   DeleteObject hBr
   DrawEdge lhDCTo, tTROut, dwFlags, BF_RECT
   hFontOld = SelectObject(lhDCTo, m_fnt.hFont)
   If (dwFlags And EDGE_SUNKEN) = EDGE_SUNKEN Then
      OffsetRect tTROut, 1, 1
   End If
   InflateRect tTROut, -2, -2
   DrawText lhDCTo, m_tBars(iButton).sCaption, -1, tTROut, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
   SelectObject lhDCTo, hFontOld
End Sub

Private Sub pPrepareMemDC(ByRef lHDC As Long, ByRef lhDCU As Long, ByRef bMemDC As Boolean)
   
   lhDCU = UserControl.hdc
   If Not m_cMemDC Is Nothing Then
      m_cMemDC.Width = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
      m_cMemDC.Height = m_iButtonHeight
      lHDC = m_cMemDC.hdc
   End If
   If lHDC = 0 Then
      lHDC = lhDCU
   Else
      bMemDC = True
   End If
   SetBkColor lHDC, TranslateColor(m_oBackColor)
   SetBkMode lHDC, TRANSPARENT
   SetTextColor lHDC, TranslateColor(m_oForeColor)

End Sub

Private Sub pMouseDown(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Dim iBtn As Long
   '
   For iBtn = 1 To m_iBarCount
      If PtInRect(m_tBars(iBtn).tR, x, y) Then
         m_iMouseDownBtn = iBtn
         pMouseMove Button, Shift, x, y
         Exit Sub
      End If
   Next iBtn
   m_iMouseDownBtn = -1
   
End Sub

Private Sub pMouseMove(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Dim iBtn As Long
Dim lHDC As Long
Dim lhDCU As Long
Dim bMemDC As Boolean
Dim iMouseOver As Long

   pPrepareMemDC lHDC, lhDCU, bMemDC
   If m_iMouseDownBtn > 0 Then
      If PtInRect(m_tBars(m_iMouseDownBtn).tR, x, y) Then
         If Not m_tBars(m_iMouseDownBtn).bPressed Then
            m_tBars(m_iMouseDownBtn).bHot = True
            m_tBars(m_iMouseDownBtn).bPressed = True
            pDrawButton lHDC, bMemDC, m_iMouseDownBtn
            pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(m_iMouseDownBtn).tR
         End If
      Else
         If m_tBars(m_iMouseDownBtn).bPressed Then
            m_tBars(m_iMouseDownBtn).bHot = False
            m_tBars(m_iMouseDownBtn).bPressed = False
            pDrawButton lHDC, bMemDC, m_iMouseDownBtn
            pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(m_iMouseDownBtn).tR
         End If
      End If
   ElseIf m_iMouseDownBtn = 0 Then
      For iBtn = 1 To m_iBarCount
         If PtInRect(m_tBars(iBtn).tR, x, y) Then
            If Not m_tBars(iBtn).bHot Then
               m_tBars(iBtn).bHot = True
               pDrawButton lHDC, bMemDC, iBtn
               pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(iBtn).tR
            End If
         Else
            If m_tBars(iBtn).bHot Then
               m_tBars(iBtn).bHot = False
               pDrawButton lHDC, bMemDC, iBtn
               pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(iBtn).tR
            End If
         End If
      Next iBtn
   End If
   
End Sub
Private Sub pSetUpBar(ByVal iBar As Long)
Dim i As Long
Dim tR As RECT
Dim lS As Long
Static s_iLastBar As Long
   
   If Len(m_sBackgroundPicture) > 0 Then
      'lvw.ForeColor = CLR_NONE
   Else
      lvw.BackColor = m_tBars(iBar).oBackColor
      lvw.ForeColor = m_tBars(iBar).oForeColor
   End If
   lvw.HighlightColor = m_tBars(iBar).oHighlightColor
   lvw.OfficeXpStyle = m_tBars(iBar).bOfficeXpStyle
   lvw.Clear
   With m_tBars(iBar)
      For i = 1 To .lItemCount
         With .tItems(i)
            lvw.Add .sCaption, .lIconIndex, , .iItemID
         End With
      Next i
   End With
   If m_tBars(iBar).eIconSize = evlbLargeIcon Then
      lvw.View = &H0 'evballvViewIcon
   Else
      lvw.View = &H2  'evballvViewSmallIcon
   End If
   lvw.Update
   pAlignLVItems tR
   lS = GetWindowLong(lvw.hWndLV, GWL_STYLE)
   lS = lS And Not WS_HSCROLL
   SetWindowLong lvw.hWndLV, GWL_STYLE, lS
   
   If Not iBar = s_iLastBar Then
      Dim cB As New cListBar
      cB.fInit m_hWndCtl, m_tBars(iBar).iID
      RaiseEvent BarClick(cB)
      
      ' Fix by Billy Propes bpropes@columbus.rr.com:
      ' Prevents the BarClick event being fired when
      ' you haven't actually changed the bar
      s_iLastBar = iBar
   End If
   
End Sub

Private Sub pMouseUp(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Dim bOver As Boolean
Dim lHDC As Long
Dim lhDCU As Long
Dim bMemDC As Boolean
Dim tR As RECT
Dim tP As POINTAPI
Dim i As Long
Dim lY As Long
Dim lW As Long
Dim lH As Long
Dim lHeight As Long
Dim lScrollHeight As Long
Dim lStep As Long
Dim bFirst As Boolean
Dim iRC As Long

   pPrepareMemDC lHDC, lhDCU, bMemDC
   '
   If m_iMouseDownBtn <> 0 Then
      If m_iMouseDownBtn > 0 Then
         bOver = PtInRect(m_tBars(m_iMouseDownBtn).tR, x, y)
         If bOver Then
            AttachMessage Me, lvw.hWndLV, WM_ERASEBKGND
            If m_iMouseDownBtn <> m_iSelBar Then
               ' Time to animate!
               GetWindowRect picScroll.hWnd, tR
               lScrollHeight = tR.Bottom - tR.Top
               lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
               If m_iMouseDownBtn < m_iSelBar Then
                  ' We want to capture the buttons below this button and the list view
                  ' to slide downwards:
                  tP.y = m_tBars(m_iMouseDownBtn + 1).tR.Top
                  ClientToScreen m_hWndCtl, tP
                  tR.Top = tP.y
               Else
                  ' we want to capture the bar and all the buttons above us to slide up:
                  tP.y = m_tBars(m_iMouseDownBtn).tR.Bottom
                  ClientToScreen m_hWndCtl, tP
                  tR.Bottom = tP.y
               End If
               m_cAnimDC.Width = tR.Right - tR.Left
               m_cAnimDC.Height = tR.Bottom - tR.Top
               BitBlt m_cAnimDC.hdc, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, lHDC, tR.Left, tR.Top, vbSrcCopy
               DeleteDC lHDC
               picScroll.Visible = False
               pSetUpBar m_iMouseDownBtn

               bFirst = True
               lStep = (tR.Bottom - tR.Top) \ 6
               lW = (tR.Right - tR.Left) * Screen.TwipsPerPixelY
               lhDCU = UserControl.hdc
               ' Make sure the control isn't too small:
               If Abs(lStep) > 0 Then
               
               If m_iMouseDownBtn < m_iSelBar Then
                  ' Scrolling down:
                  tP.x = tR.Left
                  tP.y = tR.Top
                  ScreenToClient m_hWndCtl, tP
                  lHeight = tR.Bottom - tR.Top - lStep
                  lY = tP.y * Screen.TwipsPerPixelX
                  tP.y = tP.y + lStep
                  For i = tR.Top + lStep To tR.Top + lScrollHeight Step lStep
                     BitBlt lhDCU, 0, tP.y, tR.Right - tR.Left, lHeight, m_cAnimDC.hdc, 0, 0, vbSrcCopy
                     picScroll.Move 0, lY, lW, (tP.y * Screen.TwipsPerPixelX - lY)
                     If bFirst Then
                        picLvw.Top = m_tBars(m_iMouseDownBtn).lTop
                        picScroll.Visible = True
                        bFirst = False
                     End If
                     If iRC > 4 Then
                        lvw.Update
                        iRC = 0
                     Else
                        iRC = iRC + 1
                     End If
                     tP.y = tP.y + lStep
                     lHeight = lHeight - lStep
                  Next i
               Else
                  ' Scrolling up:
                  tP.x = tR.Left
                  tP.y = tR.Top
                  ScreenToClient m_hWndCtl, tP
                  lHeight = tR.Bottom - tR.Top - lStep
                  lY = lStep
                  lH = lStep * Screen.TwipsPerPixelY
                  
                  For i = lScrollHeight + tR.Top - lStep To tR.Top Step -lStep
                     BitBlt lhDCU, 0, tP.y, tR.Right - tR.Left, lHeight, m_cAnimDC.hdc, 0, lY, vbSrcCopy
                     picScroll.Move 0, (tP.y + lHeight) * Screen.TwipsPerPixelY, lW, lH
                     If bFirst Then
                        picLvw.Top = m_tBars(m_iMouseDownBtn).lTop
                        picScroll.Visible = True
                        bFirst = False
                     End If
                     If iRC > 4 Then
                        lvw.Update
                        iRC = 0
                     Else
                        iRC = iRC + 1
                     End If
                     lHeight = lHeight - lStep
                     lY = lY + lStep
                     lH = lH + lStep * Screen.TwipsPerPixelY
                  Next i
               End If
               End If
               m_iSelBar = m_iMouseDownBtn
               Render
               picScroll.Visible = True
            End If
            DetachMessage Me, lvw.hWndLV, WM_ERASEBKGND
         End If
         
         ' Draw:
         m_tBars(m_iMouseDownBtn).bPressed = False
         m_tBars(m_iMouseDownBtn).bHot = bOver
         pDrawButton lHDC, bMemDC, m_iMouseDownBtn
         pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(m_iMouseDownBtn).tR
      End If
      m_iMouseDownBtn = 0
      
   End If
End Sub

Private Sub pScroll(ByVal iDir As Long)
Dim lY As Long
Dim lYNow As Long
Dim lStep As Long
Dim l As Long
Dim lItemHeight As Long
Dim tR As RECT
   
   lYNow = picLvw.Top
   If lvw.View = &H0 Then 'evballvViewIcon Then
      lY = lYNow + iDir * m_lItemHeight * Screen.TwipsPerPixelY
   Else
      If lvw.Count > 0 Then
         lvw.GetItemRect 1, tR.Left, tR.Top, tR.Right, tR.Bottom
      End If
      lY = lYNow + iDir * (tR.Bottom - tR.Top + 1) * Screen.TwipsPerPixelY
   End If
   
   If lY < lYNow Then
      ' We are scrolling down.  We need to move top from lYNow to lY in negative steps:
      If cmdDown.Visible Then
         lStep = (lY - lYNow) \ 4
         For l = lYNow To lY Step lStep
            picLvw.Top = l
            m_tBars(m_iSelBar).lTop = l
            lvw.Update
         Next l
      End If
   Else
      ' We are scrolling up.  We need to move top from lYNow to lY in positive steps:
      If cmdUp.Visible Then
         lStep = (lY - lYNow) \ 4
         For l = lYNow To lY Step lStep
            picLvw.Top = l
            m_tBars(m_iSelBar).lTop = l
            lvw.Update
         Next l
      End If
   End If
   
   picScroll_Resize
   If (tmrScrolling.Tag = "cmdUp" And cmdUp.Visible = False) Or (tmrScrolling.Tag = "cmdDown" And cmdDown.Visible = False) Then
      tmrScrolling.Tag = ""
      tmrScrolling.Enabled = False
   End If
   
End Sub

Private Sub Render()
Dim iBtn As Long
Dim lHDC As Long
Dim lhDCU As Long
Dim bMemDC As Boolean

   If m_iBarCount = 0 Then
      picScroll.Visible = False
   Else
      picScroll.Visible = True
      pPrepareMemDC lHDC, lhDCU, bMemDC
      For iBtn = 1 To m_iBarCount
         pDrawButton lHDC, bMemDC, iBtn
         pMemDCToDC lhDCU, lHDC, bMemDC, m_tBars(iBtn).tR
      Next iBtn
   End If
End Sub

Private Sub pMemDCToDC(ByVal lhDCU As Long, ByVal lHDC As Long, ByVal bMemDC As Boolean, ByRef tR As RECT)
   If bMemDC Then
      With tR
          BitBlt lhDCU, .Left, .Top, .Right - .Left, .Bottom - .Top, lHDC, 0, 0, vbSrcCopy
      End With
   End If
End Sub

Private Sub pAlignLVItems(tR As RECT)
Dim i As Long
Dim x As Long, y As Long
Dim lLastY As Long
Dim bAdd As Boolean
Dim tTR As RECT
Dim lOffset As Long, lOff As Long
Dim bLargeIcon As Long

   GetClientRect picScroll.hWnd, tR
   bLargeIcon = (lvw.View = &H0) ' evballvViewIcon)
   If bLargeIcon Then
      lvw.IconSpaceX = tR.Right - tR.Left - 12
      lvw.IconSpaceY = m_lItemHeight
      lvw.Update
   Else
      lvw.IconSpaceX = tR.Right - tR.Left - 4
      lvw.Update
   End If
   
   ' Now move all the lvw items so they fit:
   If lvw.Count > 0 Then
      
      If bLargeIcon Then
         ' Thanks to Rafael (rapm20@cantv.net) for noting there was
         ' a bug here.
         ' His fix worked, but I discovered this simpler one: the
         ' item position in large icon view is determined by the
         ' icon size!  Bizarre...
         lOffset = ((tR.Right - tR.Left) - 32) \ 2
      Else
         lOffset = 2
      End If
      
      For i = 1 To lvw.Count
         lvw.GetItemPosition i, x, y
         If bLargeIcon Then
            lvw.SetItemPosition i, lOffset, y
         Else
            If i > 1 Then
               lLastY = lLastY + 18
            End If
            lvw.SetItemPosition i, 4, lLastY
         End If
      Next i
               
   End If

End Sub

Private Function pbEnsureListView() As Boolean
   If m_bRunTime Then
      If lvw.hWndLV = 0 Then
         ' Only create the ListView at run time for stability!
         lvw.NoScrollBar = True
         lvw.MultiSelect = False
         lvw.HideSelection = True
         lvw.EditLabels = True
         lvw.Initialise picLvw.hWnd
         ' SPM: fixed incorrect cursor bug by moving
         ' to TwoClickActivate
         lvw.TwoClickActivate = True
         lvw.InfoTips = True
         m_hWndLV = lvw.hWndLV
         AttachMessage Me, m_hWndLV, WM_LBUTTONUP
         AttachMessage Me, m_hWndLV, WM_LBUTTONDOWN
         AttachMessage Me, m_hWndLV, WM_RBUTTONUP
         AttachMessage Me, m_hWndLV, WM_RBUTTONDOWN
         AttachMessage Me, m_hWndLV, WM_MOUSEMOVE
         AttachMessage Me, m_hWndCtl, WM_SETFOCUS
         AttachMessage Me, m_hWndLV, WM_SETFOCUS
         AttachMessage Me, m_hWndLV, WM_MOUSEACTIVATE
         AttachMessage Me, m_hWndLV, WM_MOUSEWHEEL
         
      End If
      pbEnsureListView = Not (m_hWndLV = 0)
   Else
      pbEnsureListView = True
   End If
End Function

Private Sub pInitialise()

   ' Are we a design or run control:
   m_bRunTime = UserControl.Ambient.UserMode
   ' This was only set so I could see it for design purposes:
   picScroll.BorderStyle = 0
   
   m_hWndCtl = UserControl.hWnd
   If m_bRunTime Then
      ' This property enables the other objects to get hold
      ' of me without getting into any sticky reference
      ' problems:
      SetProp m_hWndCtl, gcObjectProp, ObjPtr(Me)
      
      ' Mouse tracking:
      Set m_cTPM = New cMouseTrack
      m_cTPM.AttachMouseTracking Me
      
      ' Owner draw up/down buttons:
      Set m_cODBtn = New cOwnerDrawButton
      m_cODBtn.Attach picScroll.hWnd
      m_cODBtn.ButtonStyle(cmdUp.hWnd) = eodUp
      m_cODBtn.ButtonStyle(cmdDown.hWnd) = eodDown
   
   Else
      picLvw.Visible = False
      
   End If
   
   ' Memory DCs give us flicker free drawing ability:
   Set m_cMemDC = New cMemDC
   Set m_cAnimDC = New cMemDC
   ' Init font:
   Set m_fnt = UserControl.Font
   
   If Not (m_bRunTime) Then
      ' Some bars:
      Dim i As Long, l As Long, v As Variant
      v = Null
      For i = 1 To 2
         l = fAddBar("DT" & i, v, "Sample Bar " & i, "")
      Next i
      picScroll.BackColor = vbButtonShadow
   End If
      
End Sub
Private Sub pTerminate()
   '
   tmrScrolling.Enabled = False
   If Not m_hWndLV = 0 Then
      DetachMessage Me, m_hWndLV, WM_LBUTTONDOWN
      DetachMessage Me, m_hWndLV, WM_LBUTTONUP
      DetachMessage Me, m_hWndLV, WM_RBUTTONUP
      DetachMessage Me, m_hWndLV, WM_RBUTTONDOWN
      DetachMessage Me, m_hWndLV, WM_MOUSEMOVE
      DetachMessage Me, m_hWndCtl, WM_SETFOCUS
      DetachMessage Me, m_hWndLV, WM_SETFOCUS
      DetachMessage Me, m_hWndLV, WM_MOUSEACTIVATE
      DetachMessage Me, m_hWndLV, WM_MOUSEWHEEL
      m_hWndLV = 0
   End If
   lvw.Terminate
   Set lvw = Nothing
   RemoveProp m_hWndCtl, gcObjectProp
   Set m_cMemDC = Nothing
   Set m_cAnimDC = Nothing
   Set m_cTPM = Nothing
   If Not m_cODBtn Is Nothing Then
      m_cODBtn.ButtonStyle(cmdUp.hWnd) = eodNone
      m_cODBtn.ButtonStyle(cmdDown.hWnd) = eodNone
      m_cODBtn.Detach
   End If
   Set m_cODBtn = Nothing

End Sub

Private Function piGetShiftState() As Integer
Dim iR As Integer
Dim lR As Long
Dim lKey As Long
   iR = iR Or (-vbShiftMask * pbKeyIsPressed(vbKeyShift))
   iR = iR Or (-vbAltMask * pbKeyIsPressed(vbKeyMenu))
   iR = iR Or (-vbCtrlMask * pbKeyIsPressed(vbKeyControl))
   piGetShiftState = iR
End Function
Private Function pbKeyIsPressed( _
        ByVal nVirtKeyCode As KeyCodeConstants _
    ) As Boolean
Dim lR As Long
    lR = GetAsyncKeyState(nVirtKeyCode)
    If (lR And &H8000&) = &H8000& Then
        pbKeyIsPressed = True
    End If
End Function


Private Sub cmdDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   pScroll -1
   tmrScrolling.Tag = "cmdDown"
   tmrScrolling.Interval = 350
   tmrScrolling.Enabled = True
End Sub

Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   tmrScrolling.Tag = ""
   tmrScrolling.Enabled = False
End Sub

Private Sub cmdUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   pScroll 1
   tmrScrolling.Tag = "cmdUp"
   tmrScrolling.Interval = 350
   tmrScrolling.Enabled = True
End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   tmrScrolling.Tag = ""
   tmrScrolling.Enabled = False
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_MOUSEMOVE, WM_SETFOCUS, WM_MOUSEACTIVATE, WM_MOUSEWHEEL
      ISubclass_MsgResponse = emrPreprocess
   Case Else
      ISubclass_MsgResponse = emrConsume
   End Select
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tP As POINTAPI
Dim lItem As Long
Dim lInItem As Long
Dim lItemBefore As Long
Dim bProcessed As Boolean
Dim iButton As Integer
Dim iShift As Integer
Dim fx As Single, fy As Single
Dim tTR As RECT
Dim tLastR As RECT
   
   Select Case iMsg
   Case WM_ERASEBKGND
      ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)

   Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
      m_iDownOn = 0
      m_bDragging = False
      SetCapture lvw.hWndLV
      GetCursorPos tP
      LSet m_tDragBegin = tP
      ScreenToClient lvw.hWndLV, tP
      lItem = lvw.HitTest(tP.x, tP.y)
      If lItem > 0 Then
         If iMsg = WM_LBUTTONDOWN Then
            lvw.DownOn = lItem
            m_iDownOn = lItem
         Else
            m_iRDownOn = lItem
         End If
      Else
         ' MouseDown on ListView but not icon:
         ScreenToClient m_hWndCtl, tP
         If iMsg = WM_LBUTTONDOWN Then
            iButton = vbLeftButton
         Else
            iButton = vbRightButton
         End If
         iShift = piGetShiftState()
         fx = tP.x
         fy = tP.y
         On Error Resume Next
         fx = UserControl.ScaleX(tP.x, vbPixels, UserControl.Extender.Container.ScaleMode)
         fy = UserControl.ScaleY(tP.y, vbPixels, UserControl.Extender.Container.ScaleMode)
         On Error GoTo 0
         RaiseEvent MouseDown(iButton, iShift, fx, fy)
      End If
   
   Case WM_MOUSEMOVE
      If Not (m_iDownOn = 0) Then
         GetCursorPos tP
         If Not (m_bDragging) Then
            ' Check for drag begin:
            If Abs(m_tDragBegin.x - tP.x) > 8 Or Abs(m_tDragBegin.y - tP.y) > 8 Then
               m_bDragging = True
            End If
         Else
            ' We are dragging, check if we should offer a drop position
            ScreenToClient lvw.hWndLV, tP
            lItemBefore = -2
            lvw.GetItemRect lItem, tLastR.Left, tLastR.Top, tLastR.Right, tLastR.Bottom
            tLastR.Bottom = tLastR.Top
            For lItem = 1 To lvw.Count
               lvw.GetItemRect lItem, tTR.Left, tTR.Top, tTR.Right, tTR.Bottom
               If PtInRect(tTR, tP.x, tP.y) <> 0 Then
                  'Debug.Print "In Item "; lItem
                  lInItem = lItem
                  Exit For
               ElseIf tP.y >= tLastR.Bottom And tP.y <= tTR.Top Then
                  lItemBefore = lItem - 1
                  Exit For
               End If
               LSet tLastR = tTR
            Next lItem
            'Debug.Print lInItem, lItemBefore
            pDrawDragLine lItemBefore
         End If
      End If
      
   Case WM_LBUTTONUP, WM_RBUTTONUP
      ReleaseCapture
      m_iDownOn = 0
      m_bDragging = False
      pDrawDragLine -2
      If (iMsg = WM_LBUTTONUP And lvw.DownOn > 0) Or (iMsg = WM_RBUTTONUP And m_iRDownOn > 0) Then
         GetCursorPos tP
         ScreenToClient lvw.hWndLV, tP
         lItem = lvw.HitTest(tP.x, tP.y)
         If iMsg = WM_LBUTTONUP Then
            If lvw.DownOn > 0 Then
               If lItem = lvw.DownOn Then
                  ' That's a hit.  we raise a click event!
                  pSelectItem m_iSelBar, lItem, False
                  bProcessed = True
               End If
            End If
         Else
            If m_iRDownOn > 0 Then
               If lItem = m_iRDownOn Then
                  pSelectItem m_iSelBar, lItem, True
                  bProcessed = True
               End If
            End If
         End If
         
         lvw.DownOn = 0
         m_iRDownOn = 0
         
      End If
         
      If Not bProcessed Then
         ' MouseUp on ListView but not icon:
         ScreenToClient m_hWndCtl, tP
         If iMsg = WM_LBUTTONUP Then
            iButton = vbLeftButton
         Else
            iButton = vbRightButton
         End If
         iShift = piGetShiftState()
         fx = tP.x
         fy = tP.y
         On Error Resume Next
         fx = UserControl.ScaleX(tP.x, vbPixels, UserControl.Extender.Container.ScaleMode)
         fy = UserControl.ScaleY(tP.y, vbPixels, UserControl.Extender.Container.ScaleMode)
         On Error GoTo 0
         RaiseEvent MouseUp(iButton, iShift, fx, fy)
      End If
         
   Case WM_MOUSEWHEEL
      Dim fwKeys As Long, zDelta As Long
      fwKeys = (wParam And &HFFFF&)
      zDelta = (wParam \ &H10000)
      tP.x = (lParam And &HFFFF&)
      tP.y = (lParam \ &H10000)
      Debug.Print fwKeys, zDelta, tP.x, tP.y
   
         
   ' ------------------------------------------------------------------------------
   ' Implement focus.  Many many thanks to Mike Gainer for showing me this
   ' code.
   Case WM_SETFOCUS
      If Not lvw Is Nothing Then
         If (lvw.hWndLV = hWnd) Then
            Dim pOleObject                  As IOleObject
            Dim pOleInPlaceSite             As IOleInPlaceSite
            Dim pOleInPlaceFrame            As IOleInPlaceFrame
            Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
            Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
            Dim PosRect                     As RECT
            Dim ClipRect                    As RECT
            Dim FrameInfo                   As OLEINPLACEFRAMEINFO
            Dim grfModifiers                As Long
            Dim AcceleratorMsg              As MSG
            
            'Get in-place frame and make sure it is set to our in-between
            'implementation of IOleInPlaceActiveObject in order to catch
            'TranslateAccelerator calls
            Set pOleObject = Me
            Set pOleInPlaceSite = pOleObject.GetClientSite
            If Not pOleInPlaceSite Is Nothing Then
               pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
               If m_IPAOHookStruct.ThisPointer <> 0 Then
                  CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
                  If Not pOleInPlaceActiveObject Is Nothing Then
                     If Not pOleInPlaceFrame Is Nothing Then
                        pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
                        If Not pOleInPlaceUIWindow Is Nothing Then
                           pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
                        End If
                     End If
                  End If
                  CopyMemory pOleInPlaceActiveObject, 0&, 4
               End If
            End If
         Else
            ' THe user control:
            SetFocusAPI lvw.hWndLV
         End If
      End If
      
   Case WM_MOUSEACTIVATE
      If Not lvw Is Nothing Then
         If GetFocus() <> lvw.hWndLV Then
            SetFocusAPI Me.hWnd
            ISubclass_WindowProc = MA_NOACTIVATE
         Else
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         End If
      Else
         ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
      End If
   ' End Implement focus.
   ' ------------------------------------------------------------------------------
      
   End Select
   
End Function
Private Sub pDrawDragLine(ByVal lLine As Long)
Dim lHDC As Long
Static lLastLine As Long
Static m_tL As POINTAPI
Static lLength As Long
Dim tJunk As POINTAPI
Dim hPenOld As Long
Dim hPenBlack As Long
Dim lLeft As Long, lTop As Long
Dim lY As Long
Dim tR As RECT
Dim bErase As Boolean
   
   ' Not ready yet.
   Exit Sub
   
   lHDC = GetDC(lvw.hWndLV)
   
   SetROP2 lHDC, R2_NOTXORPEN
   hPenBlack = CreatePen(PS_SOLID, 1, &H0&)
   hPenOld = SelectObject(lHDC, hPenBlack)
   
   If lLine > 0 Then
      ' New line position:
      lvw.GetItemRect lLine, lLeft, lTop, lLength, lY
      lY = lY + 8
      GetClientRect lvw.hWndLV, tR
      tR.Left = tR.Left + 4
      tR.Right = tR.Right - 4
      If Not (m_tL.x = tR.Left And m_tL.y = lY And tR.Right - tR.Left = lLength) Then
         bErase = True
      End If
   Else
      bErase = True
   End If
   
   ' Erase last line?
   If bErase Then
      If lLastLine > 0 Then
         MoveToEx lHDC, m_tL.x, m_tL.y, tJunk
         LineTo lHDC, m_tL.x + lLength, m_tL.y
      End If
   End If
   
   If bErase Then
      If lLine > 0 Then
         m_tL.x = tR.Left
         m_tL.y = lY
         lLength = tR.Right - tR.Left
         MoveToEx lHDC, m_tL.x, m_tL.y, tJunk
         LineTo lHDC, m_tL.x + lLength, m_tL.y
      End If
   End If
   
   lLastLine = lLine
   
   SelectObject lHDC, hPenOld
   DeleteObject hPenBlack
   
   ReleaseDC lvw.hWndLV, lHDC
   
End Sub

Private Sub pSelectItem(ByVal iBar As Long, ByVal lItem As Long, Optional ByVal bRightButton As Boolean = False)
Dim tP As POINTAPI
Dim fx As Single, fy As Single

   Dim cB As New cListBar
   cB.fInit m_hWndCtl, m_tBars(m_iSelBar).iID
   Dim cI As New cListBarItem
   cI.fInit m_hWndCtl, m_tBars(m_iSelBar).iID, m_tBars(m_iSelBar).tItems(lItem).iItemID
   
   If bRightButton Then
      GetCursorPos tP
      ScreenToClient m_hWndCtl, tP
      fx = tP.x
      fy = tP.y
      On Error Resume Next
      fx = UserControl.ScaleX(tP.x, vbPixels, UserControl.Extender.Container.ScaleMode)
      fy = UserControl.ScaleY(tP.y, vbPixels, UserControl.Extender.Container.ScaleMode)
      On Error GoTo 0
      RaiseEvent ItemRightClick(cI, cB, fx, fy)
      
   Else
      m_tBars(m_iSelBar).lItemSelected = lItem
      RaiseEvent ItemClick(cI, cB)
   End If
End Sub

Private Sub lvw_CancelEdit(ByVal lIndex As Long)
   ' user has cancelled editing
   m_lEditBar = 0
   m_lEditItem = 0
   On Error Resume Next
   SetFocusAPI m_hWndCtl
End Sub

Private Sub lvw_Click()
   ' we override listview clicks
End Sub

Private Sub lvw_EndEdit(ByVal lIndex As Long, sText As String, bCancel As Boolean)
Dim cI As cListBarItem
   ' end editing.  Set the current item to sText:
   Set cI = New cListBarItem
   cI.fInit m_hWndCtl, m_tBars(m_lEditBar).iID, m_tBars(m_lEditBar).tItems(m_lEditItem).iItemID
   RaiseEvent ItemEndEdit(cI, sText, bCancel)
   If Not bCancel Then
      m_tBars(m_lEditBar).tItems(m_lEditItem).sCaption = sText
   End If
   m_lEditBar = 0
   m_lEditItem = 0
   On Error Resume Next
   SetFocusAPI m_hWndCtl
End Sub

Private Sub lvw_RequestEdit(ByVal lIndex As Long, bCancel As Boolean)
   '
End Sub

Private Sub lvw_RequestInfoTip(ByVal lIndex As Long, sInfoTip As String)
   '
   If m_iSelBar > 0 Then
      If lIndex > 0 And lIndex <= m_tBars(m_iSelBar).lItemCount Then
         If m_tBars(m_iSelBar).tItems(lIndex).sHelpText <> "" Then
            'Debug.Print m_tBars(m_iSelBar).tItems(lIndex).sHelpText
            sInfoTip = m_tBars(m_iSelBar).tItems(lIndex).sHelpText
         End If
      End If
   End If
End Sub

Private Sub m_cTPM_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   '
End Sub

Private Sub m_cTPM_MouseLeave()
   pMouseMove 0, 0, -15, -15
End Sub

Private Sub picLvw_Resize()
Dim tR As RECT
   If Not lvw Is Nothing Then
      GetClientRect picLvw.hWnd, tR
      lvw.IconSpaceX = tR.Right - tR.Left - 4
      lvw.Resize
      pAlignLVItems tR
   End If
End Sub

Private Sub picScroll_Resize()
Dim tR As RECT
Dim tIR As RECT
Dim lTop As Long
Dim lHeight As Long
Dim lDiff As Long
Dim lSize As Long
Dim bNeedUp As Boolean
Dim bNeedDown As Boolean
Dim lH As Long

On Error Resume Next

   If m_bRunTime Then
      ' Here we modify the size of the   contained LVW control:
      GetClientRect picScroll.hWnd, tR
      lHeight = (tR.Bottom - tR.Top)
      If Not (lvw.View = &H0) Then 'evballvViewIcon) Then
         If lvw.Count > 0 Then
            lvw.GetItemRect 1, tIR.Left, tIR.Top, tIR.Right, tIR.Bottom
            lSize = lvw.Count * (tIR.Bottom - tIR.Top)
         End If
      Else
         lSize = lvw.Count * m_lItemHeight
      End If
      If lSize > lHeight Then
         lHeight = lSize
         ' Scrolling needed
         lTop = picLvw.Top \ Screen.TwipsPerPixelY
         If lTop > 0 Then lTop = 0
         bNeedUp = (lTop < 0)
         bNeedDown = (lSize + lTop > tR.Bottom - tR.Top)
         lTop = lTop * Screen.TwipsPerPixelX
      Else
         ' No scrolling:
         bNeedUp = False
         bNeedDown = False
         lTop = 0
      End If
      If bNeedUp Then
         cmdUp.Move (tR.Right - 2) * Screen.TwipsPerPixelX - cmdUp.Width, (tR.Top + 2) * Screen.TwipsPerPixelY
      End If
      If bNeedUp <> cmdUp.Visible Then
         cmdUp.Visible = bNeedUp
         cmdUp.ZOrder
      End If
      If bNeedDown Then
         cmdDown.Move (tR.Right - 2) * Screen.TwipsPerPixelX - cmdUp.Width, (tR.Bottom - 2) * Screen.TwipsPerPixelY - cmdDown.Height
      End If
      If bNeedDown <> cmdDown.Visible Then
         cmdDown.Visible = bNeedDown
         cmdDown.ZOrder
      End If

      lHeight = lHeight * Screen.TwipsPerPixelY
      If (GetWindowLong(lvw.hWndLV, GWL_STYLE) And WS_HSCROLL) = WS_HSCROLL Then
         lH = GetSystemMetrics(SM_CYHSCROLL)
      End If
      lDiff = (tR.Bottom - tR.Top + lH) * Screen.TwipsPerPixelY - (lTop + lHeight)
      If lDiff > 0 Then
         lHeight = lHeight + lDiff
      End If
         
      picLvw.Move tR.Left * Screen.TwipsPerPixelX, lTop, (tR.Right - tR.Left) * Screen.TwipsPerPixelX, lHeight
   End If
End Sub

Private Sub tmrScrolling_Timer()
Dim lhWnd As Long
Dim iDir As Long
Dim tR As RECT
Dim tP As POINTAPI
   tmrScrolling.Interval = 100
   If tmrScrolling.Tag = "cmdUp" Then
      iDir = 1
      lhWnd = cmdUp.hWnd
   ElseIf tmrScrolling.Tag = "cmdDown" Then
      iDir = -1
      lhWnd = cmdDown.hWnd
   End If
   GetWindowRect lhWnd, tR
   GetCursorPos tP
   If PtInRect(tR, tP.x, tP.y) Then
      pScroll iDir
   End If
End Sub

Private Sub UserControl_Initialize()
   '
   ' Attach custom IOleInPlaceActiveObject interface
   Dim IPAO As IOleInPlaceActiveObject
   With m_IPAOHookStruct
      Set IPAO = Me
      CopyMemory .IPAOReal, IPAO, 4
      CopyMemory .TBEx, Me, 4
      .lpVTable = IPAOVTable
      .ThisPointer = VarPtr(m_IPAOHookStruct)
   End With
   
   BorderStyle = evlb3DThin
   m_iButtonHeight = 20
   m_lItemHeight = 75
   m_oBackColor = vbButtonFace
   m_oForeColor = vbWindowText
   '
   Set lvw = New cListView
   
End Sub

Private Sub UserControl_InitProperties()
   '
   pInitialise
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_bRunTime Then
      pMouseDown Button, Shift, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY
   End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_bRunTime Then
      If Not m_cTPM.Tracking Then
         m_cTPM.StartMouseTracking
      End If
      pMouseMove Button, Shift, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_bRunTime Then
      pMouseUp Button, Shift, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY
   End If
End Sub

Private Sub UserControl_Paint()
On Error Resume Next
   'If m_bRunTime Then
      Render
   'End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   '
   pInitialise
   BorderStyle = PropBag.ReadProperty("BorderStyle", evlb3DThin)
   Dim sFnt As New StdFont
   sFnt.Name = "MS Sans Serif"
   sFnt.Size = 8
   Set Font = PropBag.ReadProperty("Font", sFnt)
   BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
   ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
   ScaleMode = PropBag.ReadProperty("ScaleMode", vbTwips)
   
End Sub

Private Sub UserControl_Resize()
Static s_lWidth As Long
Dim lWidth As Long
Dim tR As RECT

On Error Resume Next
   Render
   If m_bRunTime Then
      If m_iBarCount > 0 Then
         GetWindowRect m_hWndCtl, tR
         lWidth = tR.Right - tR.Left + 1
         If Not s_lWidth = lWidth Then
            ' Reload the bar items.
            ' THis forces a resize to occur...
            pSetUpBar m_iSelBar
            s_lWidth = lWidth
         End If
      End If
   End If
End Sub

Private Sub UserControl_Terminate()
   '
   pTerminate
   
   ' Detach the custom IOleInPlaceActiveObject interface
   ' pointers.
   With m_IPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
   End With
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   '
   PropBag.WriteProperty "BorderStyle", BorderStyle, evlb3DThin
   Dim sFnt As New StdFont
   sFnt.Name = "MS Sans Serif"
   sFnt.Size = 8
   PropBag.WriteProperty "Font", Font, sFnt
   PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
   PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
   PropBag.WriteProperty "ScaleMode", ScaleMode, vbTwips
   
End Sub