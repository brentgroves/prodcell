VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmScreenSaver 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   13245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   11505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmScreenSaver.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13245
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerAlternateQualityAlerts 
      Interval        =   3000
      Left            =   120
      Top             =   0
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   10425
      Index           =   1
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   9360
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "frmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public strLoadedFile As String
Private strNewFile As String

Public iLeft As Integer
Public StartTime As Date
Public bInitialLoad As Boolean
Public rsQualityAlerts As New ADODB.Recordset
Public iInitialX As Integer
Public iInitialY As Integer
Public bFirstMouseMoveEvent As Boolean
Private iMoves As Integer
' SetCapture directs ALL mouse input to the window that has the mouse
' "captured".

Private Declare Function GetCapture& Lib "user32" ()
Private Declare Function SetCapture& Lib "user32" (ByVal hWnd&)
Private Declare Function ReleaseCapture& Lib "user32" ()





Private Sub AcroPDF_OnError(Index As Integer)
    rsQualityAlerts.Close
    Call ReleaseCapture
    Unload Me

End Sub

Private Sub Form_Click()
    rsQualityAlerts.Close
    Call ReleaseCapture
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    rsQualityAlerts.Close
    Call ReleaseCapture
    Unload Me
End Sub

Private Sub Form_Load()
'Initialize alert rotation
    On Error GoTo Reconnect
    Dim pn As Integer
    bInitialLoad = True
    bFirstMouseMoveEvent = True
    
    strLoadedFile = ""
    
    StartTime = Now
    If rsQualityAlerts.State = adStateOpen Then
        rsQualityAlerts.Close
    End If

  rsQualityAlerts.Open "Select [Document Master].filename, [Document Master].DocumentTitle " & _
 " from [Document Master] inner join [Document PartNumbers] " & _
 " on  [Document Master].DocumentId = [Document PartNumbers].DocumentId " & _
 " Where [Document Master].DocumentType = '3' and [Document PartNumbers].PartNumber = '" & MainForm.PartSelectCombo.Text & "'", SQLConn, adOpenKeyset, adLockReadOnly
    'Order by [Document PartNumbers].PartNumber 2838257
    rsQualityAlerts.MoveFirst
    strNewFile = rsQualityAlerts.Fields("FILENAME")
    
    iLeft = 30000
    TimerAlternateQualityAlerts.Enabled = True
    
    Exit Sub
'-----------------------------------------------------------------------------
Reconnect:
'-----------------------------------------------------------------------------
    Unload Me
'    MakeConnections
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    rsQualityAlerts.Close
    Call ReleaseCapture
    Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If bFirstMouseMoveEvent = True Then
    iInitialX = X
    iInitialY = Y
    bFirstMouseMoveEvent = False
  End If
  
  
'  If ((X <> iInitialX) Or (Y <> iInitialY)) Then
  If ((Y <> iInitialY)) Then
    
    rsQualityAlerts.Close
    Call ReleaseCapture
    IdleTime = Now
    Unload Me
  End If
    
End Sub

Private Sub TimerAlternateQualityAlerts_Timer()
    Dim pn As Integer
    Dim strFilePath As String
    Dim i As Integer
    Dim j As Integer
        
   Dim sec As String
   
    If bInitialLoad = True Then
      bInitialLoad = False
      Call SetCapture(frmScreenSaver.hWnd)
      frmScreenSaver.SetFocus
'              Call ReleaseCapture


    End If
    
    
    
    
    ' Calculate position to display pdf, iLeft (global)
    ' determine width of the monitor (MainForm.Width)
    ' determine width of the pdf (AcroPDF(1).width)
    ' Initialize iLeft to 0
    ' In TimerAlternateQualityAlerts_Timer() function increment iLeft by 500
    ' Check if iLeft + width of pdf is greater than monitor width.
    '   If it is set iLeft to 500
    ' disable TimerAlternateQualityAlerts if screen saver has been on 7 hours
    
    iLeft = iLeft + 500
    If ((iLeft + Me.Width) > frmBackGround.Width) Then
        iLeft = 500
        If rsQualityAlerts.EOF = True Then
            rsQualityAlerts.MoveFirst
        End If
        strNewFile = rsQualityAlerts.Fields("FILENAME")
        rsQualityAlerts.MoveNext
    End If
    
    
    Me.Left = iLeft
    
    
    If strLoadedFile <> strNewFile Then
       If 1 = SITE Then  ' Hartselle
           strFilePath = "\\hartselle-public\documentstorage\" + Trim(strNewFile)
       Else
           strFilePath = "\\busche-sql\documentstorage\" + Trim(strNewFile)
       End If
       AcroPDF(1).Visible = True
     
'    AcroPDF(1).setView "Fit"
'    AcroPDF(1).setLayoutMode "LandScape"
'    AcroPDF(1).setLayoutMode "OneColumn"
'    AcroPDF(0).setZoom (Acrobat0Zoom)
'    AcroPDF(1).setShowToolbar (False)
'    AcroPDF(1).setShowScrollbars (False)
       
    AcroPDF(1).LoadFile strFilePath
       
    AcroPDF(1).setShowToolbar (False)
    AcroPDF(1).setShowScrollbars (False)
       
    strLoadedFile = strNewFile
    End If
    
   sec = DateDiff("s", StartTime, Now)
    
    If DateDiff("h", StartTime, Now) > 7 Then
        TimerAlternateQualityAlerts.Enabled = False
        AcroPDF(1).Visible = False
    End If
    
End Sub
