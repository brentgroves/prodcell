VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FE1D1F8B-EC4B-11D3-B06C-00500427A693}#1.1#0"; "vbalLBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form MainForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   15855
   ClientLeft      =   105
   ClientTop       =   -450
   ClientWidth     =   28800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   15850
   ScaleMode       =   0  'User
   ScaleWidth      =   28805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerScreenSaver 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5160
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   7080
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Interval        =   20000
      Left            =   6120
      Top             =   840
   End
   Begin VB.CommandButton DocumentList 
      Caption         =   "Document List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   36
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton RefreshCMD 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   475
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   5160
      Top             =   840
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   2
      Left            =   20955
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   14655
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   2
      Left            =   20955
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":05A0
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   13725
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   1
      Left            =   25770
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":0B4C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   14655
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   1
      Left            =   25770
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   13725
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   0
      Left            =   13410
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":1698
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   14580
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   0
      Left            =   13410
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":1C38
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   13650
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   2
      Left            =   19380
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":21E4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   14100
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   2
      Left            =   18660
      Picture         =   "Main Form.frx":25CC
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   2
      Left            =   20115
      Picture         =   "Main Form.frx":29C1
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   2
      Left            =   17700
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":2DBE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   2
      Left            =   16260
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":31E3
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   2
      Left            =   16980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":3606
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   14700
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   2
      Left            =   16980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":3A40
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   13515
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   1
      Left            =   24165
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":3E73
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   14130
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   1
      Left            =   23445
      Picture         =   "Main Form.frx":425B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   1
      Left            =   24900
      Picture         =   "Main Form.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   1
      Left            =   22485
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":4A4D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   1
      Left            =   21045
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":4E72
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   1
      Left            =   21765
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5295
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   14730
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   1
      Left            =   21765
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":56CF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   13530
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   4125
      Left            =   15120
      TabIndex        =   11
      Top             =   1380
      Width           =   2970
      _cx             =   5239
      _cy             =   7276
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   0
      Left            =   9375
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5B02
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   13515
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   0
      Left            =   9375
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5F35
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   14700
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   0
      Left            =   8655
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":636F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   0
      Left            =   10095
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":6792
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   0
      Left            =   12510
      Picture         =   "Main Form.frx":6BB7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   0
      Left            =   11055
      Picture         =   "Main Form.frx":6FB4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   0
      Left            =   11775
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":73A9
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   14100
      Width           =   735
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   0
      Left            =   16080
      TabIndex        =   3
      Top             =   1800
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   16
      Size            =   13236
      Images          =   "Main Form.frx":7791
      Version         =   131072
      KeyCount        =   3
      Keys            =   "ÿÿ"
   End
   Begin vbalLbar6.vbalListBar ListBar 
      Height          =   14300
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   25215
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox PartSelectCombo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Main Form.frx":AB65
      Left            =   960
      List            =   "Main Form.frx":AB6C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   15360
      Width           =   28800
      _ExtentX        =   50800
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   48181
            TextSave        =   "8:39 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "5/9/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   1
      Left            =   9540
      TabIndex        =   12
      Top             =   4455
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   2
      Left            =   14400
      TabIndex        =   33
      Top             =   2760
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer Crviewer2 
      Height          =   4125
      Left            =   11160
      TabIndex        =   37
      Top             =   480
      Width           =   2970
      _cx             =   5239
      _cy             =   7276
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   13215
      Left            =   5160
      TabIndex        =   34
      Top             =   1800
      Width           =   3375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5953
      _cy             =   23310
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KeyToSend As String
Public rsScreenSaver As New ADODB.Recordset

Public bT1Enable As Boolean
Public bT2Enable As Boolean
Public bT3Enable As Boolean
Public bT4Enable As Boolean
Public bT5Enable As Boolean

Const LeftViewer = 0
Const RightViewer = 1
Const LargeViewer = 2

Private Sub RefreshCMD_Click()
    On Error GoTo Reconnect
    Dim pn As Integer
    
    pn = Me.PartSelectCombo.ListIndex
    MainForm.ListBar.Bars.Clear
    PopulateCategories
    Dim sqlrs As ADODB.Recordset
    Set sqlrs = New ADODB.Recordset
    MainForm.PartSelectCombo.Clear
    
    sqlrs.Open "SELECT DISTINCT PARTNUMBER FROM [DOCUMENT PARTNUMBERS] ORDER BY PARTNUMBER ASC", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        MainForm.PartSelectCombo.AddItem Trim(sqlrs.Fields("PARTNUMBER"))
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    Set sqlrs = Nothing
    IdleTime = Now
    MainForm.PartSelectCombo.ListIndex = pn
    PartSelectCombo_Click
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Private Sub DocumentList_Click()
    If Timer5.Enabled Then
        Exit Sub
    End If
    Timer5.Enabled = True
On Error GoTo Reconnect
    ResetViewers
    ShowLeftControls
    HideRightControls
    HideCenterControls
    DocListView = True
    craxReport2.DiscardSavedData
    craxReport2.ParameterFields.GetItemByName("PartNumber").ClearCurrentValueAndRange
    craxReport2.ParameterFields.GetItemByName("PartNumber").AddCurrentValue PartSelectCombo.Text
    MainForm.Crviewer2.ViewReport
    MainForm.Crviewer2.Refresh
    MainForm.Crviewer2.Zoom 80
    CrystalZoom2 = 80
    MainForm.Crviewer2.Left = 5055
    MainForm.Crviewer2.Visible = True
    IdleTime = Now
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Private Sub Form_Load()
    IdleTime = Now
    Init
    TimerScreenSaver.Enabled = True
    IdleTime = Now
End Sub

Private Sub LargeComboBtn_Click()
    Cl.ShowDropDownCombo PartSelectCombo
    IdleTime = Now
End Sub

'*****************************************************************************
Private Sub ListBar_ItemClick(Item As vbalLbar6.cListBarItem, Bar As vbalLbar6.cListBar)
'   ARGUMENTS:
'     RETURNS:
'   CALLED BY:
'       CALLS:
' DESCRIPTION:
'*****************************************************************************
    If Timer5.Enabled Then
        Exit Sub
    End If
    Timer5.Enabled = True
    On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT TYPE] WHERE RTRIM(LTRIM(DOCUMENTDESC)) LIKE '" + Trim(Left(Trim(Bar.Caption), InStr(Trim(Bar.Caption), "(") - 1)) + "'", SQLConn, adOpenKeyset, adLockReadOnly
        
'                                                                                                                       PCC cac001 11-23-09
'x      If sqlrs.RecordCount < 1 Then
'x              ViewDocument Str(Val(Right(Item.Key, Len(Item.Key) - 1))), Item.IconIndex, False
'x      Else
'x              ViewDocument Str(Val(Right(Item.Key, Len(Item.Key) - 1))), Item.IconIndex, sqlrs.Fields("LARGEFORMAT")
'x      End If

        '       NOTE:   Need to check file type and set ViewerType (Item.IconIndex) for pdf/mpeg
        '                       This is currently handled in ViewDocument, but should be done somewhere else - Chuck Collatz

        '=========================================================================
        '   Send the whole filename (extension striped in function "ViewDocument")
        '=========================================================================
        If (sqlrs.RecordCount < 1) Then
                ViewDocument Item.Key, Item.IconIndex, False
        Else
                ViewDocument Item.Key, Item.IconIndex, sqlrs.Fields("LARGEFORMAT")
        End If
    
    IdleTime = Now
    Exit Sub

'-----------------------------------------------------------------------------
Reconnect:
'-----------------------------------------------------------------------------
    ReconnectForm.Show
    MakeConnections
End Sub 'ListBar_ItemClick

Private Sub PageDownBtn_Click(Index As Integer)
    KeyToSend = ("{PGDN}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PageUpBtn_Click(Index As Integer)
    KeyToSend = ("{PGUP}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanDownBtn_Click(Index As Integer)
    KeyToSend = ("{DOWN}")
        If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanLeftBtn_Click(Index As Integer)
    KeyToSend = ("{LEFT}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanRightBtn_Click(Index As Integer)
    KeyToSend = ("{RIGHT}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanUpBtn_Click(Index As Integer)
    KeyToSend = ("{UP}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PartSelectCombo_Click()
    ClearDocuments
    PopulateDocuments (Trim(PartSelectCombo.Text))
    DocumentList.Enabled = True
    IdleTime = Now
End Sub

Private Sub Timer1_Timer()
    SendKeys (KeyToSend)
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    SendKeys ("^h")
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    If IdleTime < DateAdd("n", -20, Now) Then
        RefreshCMD_Click
        IdleTime = Now
    End If
End Sub

Private Sub Timer5_Timer()
    Timer5.Enabled = False
End Sub

Private Sub TimerScreenSaver_Timer()
'Dim dt As Date
 On Error GoTo Reconnect
 Dim pn As Integer
 

 If rsScreenSaver.State = adStateOpen Then
    rsScreenSaver.Close
 End If


 rsScreenSaver.Open "Select [Document Master].filename, [Document Master].DocumentTitle " & _
 " from [Document Master] inner join [Document PartNumbers] " & _
 " on  [Document Master].DocumentId = [Document PartNumbers].DocumentId " & _
 " Where [Document Master].DocumentType = '3' and [Document PartNumbers].PartNumber = '" & MainForm.PartSelectCombo.Text & "'", SQLConn, adOpenKeyset, adLockReadOnly
    'Order by [Document PartNumbers].PartNumber 2838257
 
 If ((rsScreenSaver.RecordCount > 0) And DateDiff("n", IdleTime, Now) > 15) And (MainForm.PartSelectCombo.ListIndex <> -1) Then
   TimerScreenSaver.Enabled = False
    rsScreenSaver.Close
   
    bT1Enable = MainForm.Timer1.Enabled
    bT2Enable = MainForm.Timer2.Enabled
    bT3Enable = MainForm.Timer3.Enabled
    bT4Enable = MainForm.Timer4.Enabled
    bT5Enable = MainForm.Timer5.Enabled
    
    MainForm.Timer1.Enabled = False
    MainForm.Timer2.Enabled = False
    MainForm.Timer3.Enabled = False
    MainForm.Timer4.Enabled = False
    MainForm.Timer5.Enabled = False

   frmBackGround.Show vbModal, Me
    
    MainForm.Timer1.Enabled = bT1Enable
    MainForm.Timer2.Enabled = bT2Enable
    MainForm.Timer3.Enabled = bT3Enable
    MainForm.Timer4.Enabled = bT4Enable
    MainForm.Timer5.Enabled = bT5Enable
    IdleTime = Now
    TimerScreenSaver.Enabled = True
Else
    rsScreenSaver.Close
 End If


 
 Exit Sub
 
 

'-----------------------------------------------------------------------------
Reconnect:
'-----------------------------------------------------------------------------
    ReconnectForm.Show
    MakeConnections


 'if Now - IdleTime
End Sub

Private Sub ZoomInBtn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = CrystalZoom2 + 15
        Me.Crviewer2.Zoom (CrystalZoom2)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom + 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom + 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom + 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom + 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom + 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom + 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case LargeViewer
        AcrobatLargeZoom = AcrobatLargeZoom + 15
        Me.AcroPDF(2).setZoom (AcrobatLargeZoom)
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub ZoomOutBtn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = CrystalZoom2 - 15
        Me.Crviewer2.Zoom (CrystalZoom2)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom - 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom - 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom - 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom - 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom - 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom - 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case LargeViewer
        AcrobatLargeZoom = AcrobatLargeZoom - 15
        Me.AcroPDF(2).setZoom (AcrobatLargeZoom)
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub ZoomReturn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = 80
        Me.Crviewer2.Zoom (80)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).setLayoutMode "OneColumn"
            Me.AcroPDF(0).setView "Fit"
        Case "acropdf1"
            Me.AcroPDF(1).setLayoutMode "OneColumn"
            Me.AcroPDF(1).setView "Fit"
        Case "crviewer1"
            CrystalZoom = 80
            Me.CRViewer1.Zoom (80)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).setLayoutMode "OneColumn"
            Me.AcroPDF(0).setView "Fit"
        Case "acropdf1"
            Me.AcroPDF(1).setLayoutMode "OneColumn"
            Me.AcroPDF(1).setView "Fit"
        Case "crviewer1"
            CrystalZoom = 80
            Me.CRViewer1.Zoom (80)
        End Select
    Case LargeViewer
            Me.AcroPDF(2).setLayoutMode "OneColumn"
            Me.AcroPDF(2).setView "Fit"
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub Timer4_Timer()
    Timer4.Enabled = False
    DoEvents
    MakeConnections
End Sub

