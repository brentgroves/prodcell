VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lID As Long

' only valid subsequent to pbVerify call
Private m_lPosition As Long

Friend Function fInit(ByVal hWnd As Long, ByVal lID As Long)
   m_hWnd = hWnd
   m_lID = lID
End Function

Private Function pbVerify(ByRef ctlThis As vbalListBar) As Boolean
Dim lPtr As Long
   m_lPosition = 0
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         m_lPosition = ctlThis.fIsBar(m_lID)
         If Not (m_lPosition = 0) Then
            pbVerify = True
         Else
            gErr 2, "cListBar" ' Bar does not exist
         End If
      Else
         gErr 1, "cListBar" ' object has been unloaded
      End If
   Else
      gErr 1, "cListBar" ' object has been unloaded
   End If
End Function
Public Property Get OfficeXpStyle() As Boolean
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      OfficeXpStyle = ctl.fBarOfficeXpStyle(m_lPosition)
   End If
End Property
Public Property Let OfficeXpStyle(ByVal bOfficeXpStyle As Boolean)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarOfficeXpStyle(m_lPosition) = bOfficeXpStyle
   End If
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gets/sets the bar's caption."
Attribute Caption.VB_UserMemId = 0
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Caption = ctl.fBarCaption(m_lPosition)
   End If
End Property
Public Property Let Caption(ByVal sCaption As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarCaption(m_lPosition) = sCaption
   End If
End Property
Public Property Get HelpText() As String
Attribute HelpText.VB_Description = "Gets/sets the tooltip that will be shown when the mouse hovers over the bar's selection button."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ' Fix by Simon Horton (simon@capio.com)
      ' Was setting the caption rather than
      ' returning the HelpText....
      HelpText = ctl.fBarHelpText(m_lPosition)
   End If
End Property
Public Property Let HelpText(ByVal sHelpText As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarHelpText(m_lPosition) = sHelpText
   End If
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Gets/sets the key of the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Key = ctl.fBarKey(m_lPosition)
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarKey(m_lPosition) = sKey
   End If
End Property
Public Property Get Tag() As String
Attribute Tag.VB_Description = "Gets/sets a string tag which can be used to associated data with the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Tag = ctl.fBarTag(m_lPosition)
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarTag(m_lPosition) = sTag
   End If
End Property
Public Property Get Index() As Long
Attribute Index.VB_Description = "Gets the 1-based numeric index of the bar in the ListBars collection."
   Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Index = m_lPosition
   End If
End Property
Public Property Get Items() As cListBarItems
Attribute Items.VB_Description = "Gets a reference to the collection of items within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Dim cLI As New cListBarItems
      cLI.fInit m_hWnd, m_lID
      Set Items = cLI
   End If
End Property
Public Property Get SelectedItem() As cListBarItem
Attribute SelectedItem.VB_Description = "Gets the last selected item within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Dim cLI As New cListBarItem
      cLI.fInit m_hWnd, m_lID, ctl.fBarSelectedItemID(m_lPosition)
      Set SelectedItem = cLI
   End If
End Property
Public Property Get IconSize() As EVBALLBIconSizeConstants
Attribute IconSize.VB_Description = "Gets/sets the size of the icons in the Bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      IconSize = ctl.fBarIconSize(m_lPosition)
   End If
End Property
Public Property Let IconSize(ByVal eSize As EVBALLBIconSizeConstants)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarIconSize(m_lPosition) = eSize
   End If
End Property
Public Property Get HighlightColor() As OLE_COLOR
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      HighlightColor = ctl.fBarHighlightColor(m_lPosition)
   End If
End Property
Public Property Let HighlightColor(ByVal oColor As OLE_COLOR)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarHighlightColor(m_lPosition) = oColor
   End If
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the background color of the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      BackColor = ctl.fBarBackColor(m_lPosition)
   End If
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarBackColor(m_lPosition) = oColor
   End If
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the foreground (text) color of the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ' Fix by Simon Horton (simon@capio.com)
      ' Was returning the bar's back colour and not
      ' the fore colour...
      ForeColor = ctl.fBarForeColor(m_lPosition)
   End If
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarForeColor(m_lPosition) = oColor
   End If
End Property
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Gets whether the bar is currently selected within the control."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Selected = ctl.fBarSelected(m_lPosition)
   End If
End Property
Public Sub SelectBar()
Attribute SelectBar.VB_Description = "Selects the bar within the control if it is not already selected."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarSelect m_lPosition
   End If
End Sub

Public Sub Sort(ByVal eDir As EVBALLBBSortOrderConstants)
Attribute Sort.VB_Description = "Sorts the items in the bar according to their caption."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarSort m_lPosition, eDir
   End If
End Sub