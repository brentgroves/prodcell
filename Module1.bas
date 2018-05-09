Attribute VB_Name = "Module1"
Public SQLConn As ADODB.Connection
Public SQLConnTOOLLIST As ADODB.Connection
Public Const Acrobat = 2
Public Const Crystal = 1
Public Const MPEG = 3
Public craxReport As New CRAXDRT.Report
Public craxApp As New CRAXDRT.Application
Public craxReport2 As New CRAXDRT.Report
Public craxApp2 As New CRAXDRT.Application
Public RightView As String
Public LeftView As String
Public CrystalZoom As Integer
Public CrystalZoom2 As Integer
Public Acrobat0Zoom As Integer
Public Acrobat1Zoom As Integer
Public AcrobatLargeZoom As Integer
Public IdleTime As Date
Public DocListView As Boolean

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As String) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SITE = 0  ' 0 for Indiana, 1 for Alabama
Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40
Global Const res19x10 = 1
Global Const res16x10 = 0.86

' initial shiftControls only for monitor with different res
Global Const shiftControls = res16x10

Global Const accWidth = 5500 * shiftControls
Global Const lViewWidth = 11875 * shiftControls
Global Const fullScreenWidth = 23750 * shiftControls


'Global Const statusWidth = 28805 'Cant be set for this control. Change it a compile time
'Global Const panelWidth = 500 'Cant be set for this control.
'Date panel width 1440.25


' 5500 * x = 4600
'Global Const accWidth = 5500 * 0.875
'Global Const lViewWidth = 11875 * 0.875


Public Sub Init()
 '       ToggleTaskBar                                           'NOTE: Remove this for debug development
        MainForm.Top = 0
        MainForm.Left = 0
        MainForm.DocumentList.Width = accWidth
        MainForm.ListBar.Width = accWidth
        MainForm.PartSelectCombo.Width = accWidth - 975
        MainForm.PanDownBtn(0).Left = MainForm.PanDownBtn(0).Left * shiftControls
        MainForm.PanLeftBtn(0).Left = MainForm.PanLeftBtn(0).Left * shiftControls
        MainForm.PanUpBtn(0).Left = MainForm.PanUpBtn(0).Left * shiftControls
        MainForm.PanRightBtn(0).Left = MainForm.PanRightBtn(0).Left * shiftControls
        MainForm.ZoomInBtn(0).Left = MainForm.ZoomInBtn(0).Left * shiftControls
        MainForm.ZoomOutBtn(0).Left = MainForm.ZoomOutBtn(0).Left * shiftControls
        MainForm.ZoomReturn(0).Left = MainForm.ZoomReturn(0).Left * shiftControls
        MainForm.PageDownBtn(0).Left = MainForm.PageDownBtn(0).Left * shiftControls + 50
        MainForm.PageUpBtn(0).Left = MainForm.PageUpBtn(0).Left * shiftControls + 50
        
        MainForm.PanDownBtn(1).Left = MainForm.PanDownBtn(1).Left * shiftControls
        MainForm.PanLeftBtn(1).Left = MainForm.PanLeftBtn(1).Left * shiftControls
        MainForm.PanUpBtn(1).Left = MainForm.PanUpBtn(1).Left * shiftControls
        MainForm.PanRightBtn(1).Left = MainForm.PanRightBtn(1).Left * shiftControls
        MainForm.ZoomInBtn(1).Left = MainForm.ZoomInBtn(1).Left * shiftControls
        MainForm.ZoomOutBtn(1).Left = MainForm.ZoomOutBtn(1).Left * shiftControls
        MainForm.ZoomReturn(1).Left = MainForm.ZoomReturn(1).Left * shiftControls
        MainForm.PageDownBtn(1).Left = MainForm.PageDownBtn(1).Left * shiftControls + 50
        MainForm.PageUpBtn(1).Left = MainForm.PageUpBtn(1).Left * shiftControls + 50
        
        MainForm.PanDownBtn(2).Left = MainForm.PanDownBtn(2).Left * shiftControls
        MainForm.PanLeftBtn(2).Left = MainForm.PanLeftBtn(2).Left * shiftControls
        MainForm.PanUpBtn(2).Left = MainForm.PanUpBtn(2).Left * shiftControls
        MainForm.PanRightBtn(2).Left = MainForm.PanRightBtn(2).Left * shiftControls
        MainForm.ZoomInBtn(2).Left = MainForm.ZoomInBtn(2).Left * shiftControls
        MainForm.ZoomOutBtn(2).Left = MainForm.ZoomOutBtn(2).Left * shiftControls
        MainForm.ZoomReturn(2).Left = MainForm.ZoomReturn(2).Left * shiftControls
        MainForm.PageDownBtn(2).Left = MainForm.PageDownBtn(2).Left * shiftControls + 50
        MainForm.PageUpBtn(2).Left = MainForm.PageUpBtn(2).Left * shiftControls + 50
        
        
'        MainForm.StatusBar1.Width = statusWidth 'Cant be set for this control. Change it a compile time
        
        
        
        
        
'        MainForm.Width = 28850
'        MainForm.Height = 15850
' Panen Width Cant be set like this. Check Dnc version to manipulate panel or scrap scan program
 '       MainForm.StatusBar1.Panels(1).Text = "Panel 1"
 '       MainForm.StatusBar1.Panels(2).Text = "Panel 2"
 '       MainForm.StatusBar1.Panels(1).Width = panelWidth
  '      MainForm.StatusBar1.Panels(2).Width = panelWidth
'        MainForm.StatusBar1.Panels(1).Width = 1000
 '       MainForm.StatusBar1.Panels(2).Width = 2000
        
        
        MainForm.ListBar.ImageList = MainForm.vbalImageList1
        MakeConnections
        SetOriginalViewerPositions
        HideLeftControls
        HideRightControls
        HideCenterControls
End Sub

Public Sub MakeConnections()
    On Error GoTo Retry
    Set SQLConn = New ADODB.Connection
    Set SQLConnTOOLLIST = New ADODB.Connection
    
    If 1 = SITE Then  ' Hartselle
        SQLConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=busche document management;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
        SQLConnTOOLLIST.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=BUSCHE TOOLLIST;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
    Else ' Indiana
        SQLConn.Open "Provider=sqloledb;" & _
           "Data Source=busche-sql;" & _
           "Initial Catalog=busche document management;" & _
           "User Id=sa;" & _
           "Password=buschecnc1"
        SQLConnTOOLLIST.Open "Provider=sqloledb;" & _
               "Data Source=busche-sql;" & _
               "Initial Catalog=BUSCHE TOOLLIST;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
    End If
    
    
    LoadPartNumbers
    
    ClearDocuments
    InitializeReport
    ReconnectForm.Hide
    Exit Sub
Retry:
    MainForm.Timer4.Enabled = True
End Sub

Public Sub LoadPartNumbers()
On Error GoTo Reconnect
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
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Private Sub ToggleTaskBar()
    Dim TaskBarWnd As Long
    TaskBarWnd = FindWindow("Shell_TrayWnd", vbNullString)
    If IsWindowVisible(TaskBarWnd) Then
        Call ShowWindow(TaskBarWnd, SW_HIDE)
    End If
End Sub

Public Sub PopulateDocuments(PartNumber As String)
On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Dim TEMP
    Dim IsPDF As Long
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT MASTER] INNER JOIN [DOCUMENT PARTNUMBERS] ON [DOCUMENT MASTER].DOCUMENTID = [DOCUMENT PARTNUMBERS].DOCUMENTID WHERE [PARTNUMBER] = '" + PartNumber + "' AND ACTIVE = 1 AND GLOBALDOC = 0 ORDER BY DOCUMENTTITLE", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        If Right(sqlrs.Fields("FILENAME"), 3) = "pdf" Then
            IsPDF = 3
        Else
            IsPDF = 2
        End If
        TEMP = MainForm.ListBar.Bars("A" + Trim(Str(sqlrs.Fields("DOCUMENTTYPE")))).Items.Add("A" + Trim(sqlrs.Fields("FILENAME")), , Trim(sqlrs.Fields("DOCUMENTTITLE")), IsPDF)
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    sqlrs.Open "SELECT * FROM [DOCUMENT MASTER] WHERE ACTIVE = 1 AND GLOBALDOC = 1 ORDER BY DOCUMENTTITLE", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        If Right(sqlrs.Fields("FILENAME"), 3) = "pdf" Then
            IsPDF = 3
        Else
            IsPDF = 2
        End If
        TEMP = MainForm.ListBar.Bars("A" + Trim(Str(sqlrs.Fields("DOCUMENTTYPE")))).Items.Add("A" + Trim(sqlrs.Fields("FILENAME")), , Trim(sqlrs.Fields("DOCUMENTTITLE")), IsPDF)
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    sqlrs.Open "SELECT * FROM [TOOLLIST MASTER] INNER JOIN [TOOLLIST PARTNUMBERS] ON [TOOLLIST MASTER].PROCESSID = [TOOLLIST PARTNUMBERS].PROCESSID WHERE [PARTNUMBERS] = '" + PartNumber + "' AND (([REVOFPROCESSID] = 0 AND [REVINPROCESS] = 0) OR ([REVOFPROCESSID] <> 0 AND [REVINPROCESS] <> 0) OR ([REVOFPROCESSID] = 0 AND [REVINPROCESS] <> 0))", SQLConnTOOLLIST, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        TEMP = MainForm.ListBar.Bars("TOOLLIST").Items.Add("A" + Trim(Str(sqlrs.Fields("PROCESSID"))), , Trim(sqlrs.Fields("OPERATIONDESCRIPTION")), 1)
  
        sqlrs.MoveNext
    Wend
    Set sqlrs = Nothing
    Dim i
    i = 0
    MainForm.ListBar.Bars("TOOLLIST").Caption = MainForm.ListBar.Bars("TOOLLIST").Caption + "  (" + Trim(Str(MainForm.ListBar.Bars("TOOLLIST").Items.Count)) + " Docs)"
    For i = 0 To 200
        On Error Resume Next
        MainForm.ListBar.Bars("A" + Trim(Str(i))).Caption = MainForm.ListBar.Bars("A" + Trim(Str(i))).Caption + "  (" + Trim(Str(MainForm.ListBar.Bars("A" + Trim(Str(i))).Items.Count)) + " Docs)"
    Next
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub ClearDocuments()
    MainForm.ListBar.Bars.Clear
    PopulateCategories
End Sub

'*****************************************************************************
'       NOTE:   Changed the 1st parameter type                                  PCC cac001 11-23-09
'ublic Sub ViewDocument(DocumentID As Integer, ViewerType As Integer, Landscape As Boolean)
Public Sub ViewDocument(DocumentID As String, ViewerType As Integer, Landscape As Boolean)
'   ARGUMENTS:
'     RETURNS:
'   CALLED BY:
'       CALLS:
' DESCRIPTION:
'*****************************************************************************
    On Error GoTo Reconnect

    Dim doc_fname   As String
    ' The Tool List Crystal Report is expecting an integer parameter so convert DocumentId to string type
    Dim intDocumentID As Long
    
        '=========================================================================
        '
        '=========================================================================
    ResetViewers

'Start of additions                                                                                     PCC cac001 11-23-09
    '=========================================================================
    '   Temp fix to catch video files because ViewerType is not set
    '   Just look for the ".mpg" extension
    '
    '   NOTE:   The correct solution should be to set/use ViewerType of MPEG
    '=========================================================================
    doc_fname = Str(Val(Right(DocumentID, Len(DocumentID) - 1)))

    '                                     3 = get right 3 extension chars
    '                                         2 = convert to lower case
        If (StrComp(StrConv(Right(DocumentID, 3), 2), "mpg", vbTextCompare) = 0) Then
        '                                                                                                               =0 means strings compare OK
    
'               ResetViewers   Allready done above
                HideLeftControls
                HideRightControls
'xxx    ShowCenterControls
                HideCenterControls

                MainForm.WindowsMediaPlayer1.settings.autoStart = False
                MainForm.WindowsMediaPlayer1.stretchToFit = True
                MainForm.WindowsMediaPlayer1.Visible = True

                MainForm.WindowsMediaPlayer1.Left = accWidth
                MainForm.WindowsMediaPlayer1.Top = 0
            '    MainForm.WindowsMediaPlayer1.Height = 13400
                MainForm.WindowsMediaPlayer1.Height = 15335
'                MainForm.WindowsMediaPlayer1.Width = 23745
                MainForm.WindowsMediaPlayer1.Width = fullScreenWidth - 5

  '              MainForm.WindowsMediaPlayer1.Left = 5060
   '             MainForm.WindowsMediaPlayer1.Top = 0
    '        '    MainForm.WindowsMediaPlayer1.Height = 13400
    '            MainForm.WindowsMediaPlayer1.Height = 15335
     '           MainForm.WindowsMediaPlayer1.Width = 23745


                If 1 = SITE Then  ' Hartselle
                    MainForm.WindowsMediaPlayer1.URL = "\\hartselle-public\documentstorage\" + Trim(doc_fname) + ".mpg"
                Else
                    MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(doc_fname) + ".mpg"
                End If


                RightView = ""

                Exit Sub
    End If
'End of additions                                                                                       PCC cac001 11-23-09

    '=========================================================================
    '
    '=========================================================================
    If craxReport.ReportTitle <> "Busche Tool List" Then
        
        If 1 = SITE Then  ' Hartselle
            Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
        Else
            Set craxReport = craxApp.OpenReport("\\buschesv2\public\Report Files\toollist.rpt")
        End If
        
        craxReport.DiscardSavedData
        craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
        craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (0)
        MainForm.CRViewer1.ReportSource = craxReport
        MainForm.CRViewer1.Zoom 80
        CrystalZoom = 80
    End If
    
    '=========================================================================
    '
    '=========================================================================
    If Landscape = True Then
        HideLeftControls
        HideRightControls
        ShowCenterControls
        If ViewerType = Acrobat Then
            If 1 = SITE Then  ' Hartselle
                MainForm.AcroPDF(2).LoadFile "\\hartselle-public\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
            Else
                MainForm.AcroPDF(2).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
            End If
            
            MainForm.AcroPDF(2).Visible = True
            MainForm.AcroPDF(2).setShowScrollbars (True)
            HideNavigationPanel (2)
        ElseIf ViewerType = MPEG Then
            'XXX Should not get here! (until ViewerType is set correctly) - Chuck Collatz
            MainForm.WindowsMediaPlayer1.Visible = False 'True
'           MainForm.WindowsMediaPlayer1.URL = Trim(Str(documenid)) + ".mpg"
'xxx        MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".mpg"
            'MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".mpg"
            'MainForm.WindowsMediaPlayer1.play
            'MainForm.WindowsMediaPlayer1.stretchToFit = True
            RightView = ""
        End If
    Else
        Select Case LCase(Trim(RightView))
'Global Const accWidth = 5500
'Global Const lViewWidth = 11875
            Case "acropdf0"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(1).Left = accWidth + lViewWidth
'                    MainForm.AcroPDF(1).Left = 5055 + 11875
                    MainForm.AcroPDF(1).Visible = True
                   If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(1).LoadFile "\\hartselle-public\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(1).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(1).setShowToolbar (False)
                    MainForm.AcroPDF(1).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "1"
                    MainForm.AcroPDF(0).Left = accWidth
'                    MainForm.AcroPDF(0).Left = 5055
                    MainForm.AcroPDF(0).Visible = True
                    HideNavigationPanel (1)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = accWidth + lViewWidth
'                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)

                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    LeftView = RightView
                    RightView = MainForm.CRViewer1.Name
'                    MainForm.AcroPDF(0).Left = 5055
                    MainForm.AcroPDF(0).Left = accWidth
                    MainForm.AcroPDF(0).Visible = True
                End If
                                                        '-------------------------------------------------
            Case "acropdf1"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Left = accWidth + lViewWidth
                    MainForm.AcroPDF(0).Visible = True
                   If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    MainForm.AcroPDF(1).Left = accWidth
                    MainForm.AcroPDF(1).Visible = True
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = accWidth + lViewWidth
'                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    LeftView = RightView
                    RightView = MainForm.CRViewer1.Name
                    MainForm.AcroPDF(1).Left = accWidth
'                    MainForm.AcroPDF(1).Left = 5055
                    MainForm.AcroPDF(1).Visible = True
                End If
                                                        '-------------------------------------------------
            Case "crviewer1"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Left = accWidth + lViewWidth
                 '   MainForm.AcroPDF(0).Left = 5055 + 11875
                    MainForm.AcroPDF(0).Visible = True
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    MainForm.CRViewer1.Left = accWidth
'                    MainForm.CRViewer1.Left = 5055
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = accWidth + lViewWidth
'                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    RightView = MainForm.CRViewer1.Name
                    HideLeftControls
                End If
                                                        '-------------------------------------------------
            Case ""
                                                        '-------------------------------------------------
                HideCenterControls
                HideLeftControls
                ShowRightControls
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Visible = True
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = accWidth + lViewWidth
'                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    RightView = MainForm.CRViewer1.Name
                End If
            End Select
     End If
     
     Exit Sub
     
'=============================================================================
Reconnect:
'=============================================================================
    ReconnectForm.Show
    MakeConnections
End Sub 'ViewDocument

Public Sub PopulateCategories()
On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Dim i As Integer
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT TYPE] ORDER BY DOCUMENTDESC ASC", SQLConn, adOpenKeyset, adLockReadOnly
    MainForm.ListBar.Bars.Add "TOOLLIST", , "Tool List"
    While Not sqlrs.EOF
        MainForm.ListBar.Bars.Add "A" + Trim(Str(sqlrs.Fields("DocumentTypeID"))), , Trim(sqlrs.Fields("DocumentDesc"))
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    Set sqlrs = Nothing
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub InitializeReport()
On Error GoTo Reconnect
   
    If 1 = SITE Then  ' Hartselle
        Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
    Else ' Indiana
        Set craxReport = craxApp.OpenReport("\\buschesv2\public\Report Files\toollist.rpt")
    End If
    
    craxReport.DiscardSavedData
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (0)
    MainForm.CRViewer1.ReportSource = craxReport
    MainForm.CRViewer1.ViewReport
    MainForm.CRViewer1.Zoom 80
    CrystalZoom = 80

    If 1 = SITE Then  ' Hartselle
        Set craxReport2 = craxApp2.OpenReport("\\hartselle-public\Shared\Public\Report Files\Document List.rpt")
    Else
        Set craxReport2 = craxApp2.OpenReport("\\buschesv2\public\Report Files\Document List.rpt")
    End If
    
    craxReport2.DiscardSavedData
    craxReport2.ParameterFields.GetItemByName("PartNumber").ClearCurrentValueAndRange
    craxReport2.ParameterFields.GetItemByName("PartNumber").AddCurrentValue ("")
    MainForm.Crviewer2.ReportSource = craxReport2
    MainForm.Crviewer2.ViewReport
    MainForm.Crviewer2.Zoom 80
    CrystalZoom2 = 80
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub SetOriginalViewerPositions()
    MainForm.Crviewer2.Left = -12000
    MainForm.Crviewer2.Top = 0
    MainForm.Crviewer2.Height = 13400
    MainForm.Crviewer2.Width = lViewWidth
'    MainForm.Crviewer2.Width = 11875
    MainForm.CRViewer1.Left = -12000
    MainForm.CRViewer1.Top = 0
    MainForm.CRViewer1.Height = 13400
    MainForm.CRViewer1.Width = lViewWidth
'    MainForm.CRViewer1.Width = 11875
    
'Global Const accWidth = 5500 * 0.875
'Global Const lViewWidth = 11875 * 0.875
    
    MainForm.AcroPDF(0).Left = accWidth + lViewWidth
    MainForm.AcroPDF(0).Top = 0
    MainForm.AcroPDF(0).Width = lViewWidth
    MainForm.AcroPDF(0).Height = 13400
    
'    MainForm.AcroPDF(0).Left = 5055 + 11875
 '   MainForm.AcroPDF(0).Top = 0
  '  MainForm.AcroPDF(0).Width = 11875
   ' MainForm.AcroPDF(0).Height = 13400
    
    MainForm.AcroPDF(1).Left = accWidth + lViewWidth
    MainForm.AcroPDF(1).Top = 0
    MainForm.AcroPDF(1).Width = lViewWidth
    MainForm.AcroPDF(1).Height = 13400
    
'    MainForm.AcroPDF(1).Left = 5055 + 11875
 '   MainForm.AcroPDF(1).Top = 0
  '  MainForm.AcroPDF(1).Width = 11875
   ' MainForm.AcroPDF(1).Height = 13400
    
    MainForm.AcroPDF(2).Left = accWidth
    MainForm.AcroPDF(2).Top = 0
    MainForm.AcroPDF(2).Width = fullScreenWidth
    MainForm.AcroPDF(2).Height = 13400
    
'    MainForm.AcroPDF(2).Left = 5055
 '   MainForm.AcroPDF(2).Top = 0
  '  MainForm.AcroPDF(2).Width = 23750
   ' MainForm.AcroPDF(2).Height = 13400
    
    MainForm.WindowsMediaPlayer1.Left = accWidth
    MainForm.WindowsMediaPlayer1.Top = 0
    MainForm.WindowsMediaPlayer1.Height = 13400
    MainForm.WindowsMediaPlayer1.Width = fullScreenWidth
    MainForm.WindowsMediaPlayer1.Visible = False
'    MainForm.WindowsMediaPlayer1.Left = 5055
 '   MainForm.WindowsMediaPlayer1.Top = 0
  '  MainForm.WindowsMediaPlayer1.Height = 13400
   ' MainForm.WindowsMediaPlayer1.Width = 23750
    'MainForm.WindowsMediaPlayer1.Visible = False
    
    MainForm.AcroPDF(0).Visible = False
    MainForm.AcroPDF(1).Visible = False
    MainForm.AcroPDF(2).Visible = False
    AcrobatLeftZoom = 80
    AcrobatRightZoom = 80
    AcrobatLargeZoom = 100
End Sub

Public Sub ShowLeftControls()
    With MainForm
            .PanDownBtn(0).Visible = True
            .PanLeftBtn(0).Visible = True
            .PanUpBtn(0).Visible = True
            .PanRightBtn(0).Visible = True
            .ZoomInBtn(0).Visible = True
            .ZoomOutBtn(0).Visible = True
            .ZoomReturn(0).Visible = True
            .PageDownBtn(0).Visible = True
            .PageUpBtn(0).Visible = True
    End With
End Sub

Public Sub ShowRightControls()
    With MainForm
            .PanDownBtn(1).Visible = True
            .PanLeftBtn(1).Visible = True
            .PanUpBtn(1).Visible = True
            .PanRightBtn(1).Visible = True
            .ZoomInBtn(1).Visible = True
            .ZoomOutBtn(1).Visible = True
            .ZoomReturn(1).Visible = True
            .PageDownBtn(1).Visible = True
            .PageUpBtn(1).Visible = True
    End With
End Sub

Public Sub HideLeftControls()
    With MainForm
            .PanDownBtn(0).Visible = False
            .PanLeftBtn(0).Visible = False
            .PanUpBtn(0).Visible = False
            .PanRightBtn(0).Visible = False
            .ZoomInBtn(0).Visible = False
            .ZoomOutBtn(0).Visible = False
            .ZoomReturn(0).Visible = False
            .PageDownBtn(0).Visible = False
            .PageUpBtn(0).Visible = False
    End With
End Sub

Public Sub HideRightControls()
    With MainForm
            .PanDownBtn(1).Visible = False
            .PanLeftBtn(1).Visible = False
            .PanUpBtn(1).Visible = False
            .PanRightBtn(1).Visible = False
            .ZoomInBtn(1).Visible = False
            .ZoomOutBtn(1).Visible = False
            .ZoomReturn(1).Visible = False
            .PageDownBtn(1).Visible = False
            .PageUpBtn(1).Visible = False
    End With
End Sub

Public Sub ShowCenterControls()
    With MainForm
            .PanDownBtn(2).Visible = True
            .PanLeftBtn(2).Visible = True
            .PanUpBtn(2).Visible = True
            .PanRightBtn(2).Visible = True
            .ZoomInBtn(2).Visible = True
            .ZoomOutBtn(2).Visible = True
            .ZoomReturn(2).Visible = True
            .PageDownBtn(2).Visible = True
            .PageUpBtn(2).Visible = True
    End With
End Sub

Public Sub HideCenterControls()
    With MainForm
            .PanDownBtn(2).Visible = False
            .PanLeftBtn(2).Visible = False
            .PanUpBtn(2).Visible = False
            .PanRightBtn(2).Visible = False
            .ZoomInBtn(2).Visible = False
            .ZoomOutBtn(2).Visible = False
            .ZoomReturn(2).Visible = False
            .PageDownBtn(2).Visible = False
            .PageUpBtn(2).Visible = False
    End With
End Sub

Public Sub ResetViewers()
    DocListView = False
'Global Const accWidth = 5500 * 0.875
'Global Const lViewWidth = 11875 * 0.875
    
    
    MainForm.CRViewer1.Left = -12000
    MainForm.Crviewer2.Left = -12000
    MainForm.AcroPDF(0).Left = accWidth + lViewWidth
    MainForm.AcroPDF(1).Left = accWidth + lViewWidth
'    MainForm.AcroPDF(0).Left = 5055 + 11875
'    MainForm.AcroPDF(1).Left = 5055 + 11875
    
'PCC cac001 11-23-09
    '=========================================================================
        '       Move the media player out of the visiable display
    '   Since it left a shadow area where the PDF should display
        '       it is moved off the visiable area of the display when not in use
    '=========================================================================
    MainForm.WindowsMediaPlayer1.Left = accWidth
'    MainForm.WindowsMediaPlayer1.Left = 5055
    MainForm.WindowsMediaPlayer1.Top = -13500 '0
    MainForm.WindowsMediaPlayer1.Height = 13400
    MainForm.WindowsMediaPlayer1.Width = fullScreenWidth
'    MainForm.WindowsMediaPlayer1.Width = 23750
    MainForm.WindowsMediaPlayer1.Visible = False
    
    MainForm.AcroPDF(0).Visible = False
    MainForm.AcroPDF(1).Visible = False
    MainForm.AcroPDF(2).Visible = False
End Sub

Public Sub HideNavigationPanel(Viewer As Integer)
    MainForm.AcroPDF(Viewer).setView "FIT"
    MainForm.AcroPDF(Viewer).setShowToolbar (True)
    MainForm.AcroPDF(Viewer).SetFocus
    MainForm.Timer2.Enabled = True
    While MainForm.Timer2.Enabled = True
        DoEvents
    Wend
    MainForm.AcroPDF(Viewer).setShowToolbar (False)
End Sub
