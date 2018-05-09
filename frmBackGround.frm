VERSION 5.00
Begin VB.Form frmBackGround 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10290
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   12690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBackGround.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2160
      Top             =   2280
   End
End
Attribute VB_Name = "frmBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub Timer1_Timer()
On Error GoTo err1

Timer1.Enabled = False
frmScreenSaver.Show vbModal, Me
Unload Me
Exit Sub
err1:
Unload Me

End Sub
