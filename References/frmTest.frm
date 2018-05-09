VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "vbAccelerator Flat Control Class Tester"
   ClientHeight    =   6765
   ClientLeft      =   2865
   ClientTop       =   2175
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6540
   Begin VB.OptionButton optStyle 
      Caption         =   "Office &11"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   3540
      Width           =   2895
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Office 1&0"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   3300
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Office &9"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   3060
      Width           =   2895
   End
   Begin VB.PictureBox picDemo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   3360
      Picture         =   "frmTest.frx":1272
      ScaleHeight     =   660
      ScaleWidth      =   2535
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2820
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   2160
      Width           =   3060
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.ComboBox cboSize 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CheckBox chkEnabled 
      Appearance      =   0  'Flat
      Caption         =   "&Enabled"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   6360
      Value           =   1  'Checked
      Width           =   6435
   End
   Begin VB.ComboBox cboDropDown 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3075
   End
   Begin VB.ComboBox cboDisabled 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Text            =   "cboDisabled"
      Top             =   5940
      Width           =   3015
   End
   Begin VB.ComboBox cboList 
      Height          =   315
      ItemData        =   "frmTest.frx":1AFB
      Left            =   60
      List            =   "frmTest.frx":1AFD
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Frame fraTest 
      Caption         =   "Testing Controls inside frame"
      Height          =   1215
      Left            =   3240
      TabIndex        =   17
      Top             =   5040
      Width           =   3255
      Begin VB.PictureBox picFrameFixer 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   60
         ScaleHeight     =   915
         ScaleWidth      =   3135
         TabIndex        =   32
         Top             =   240
         Width           =   3135
         Begin VB.ComboBox cboInFrame 
            Height          =   315
            Left            =   0
            TabIndex        =   34
            Text            =   "Combo1"
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtInFrame 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "frmTest.frx":1AFF
            Top             =   60
            Width           =   3000
         End
      End
   End
   Begin VB.PictureBox picBorder 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1215
      ScaleWidth      =   3255
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3255
      Begin VB.ComboBox cboInPicture 
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   780
         Width           =   3015
      End
      Begin VB.TextBox txtInPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label lblPic 
         Caption         =   "Testing Controls Inside Picture Box"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   3015
      End
   End
   Begin VB.ComboBox cboTest 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Text            =   "cboTest"
      Top             =   5100
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmTest.frx":1B05
      Top             =   4140
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   120
      Picture         =   "frmTest.frx":1B0B
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H00000066&
      Height          =   435
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1920
      TabIndex        =   27
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   1680
      TabIndex        =   26
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   25
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H0080C0FF&
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   24
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H00C0C000&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   23
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblColorBit 
      BackColor       =   &H000080FF&
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   22
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblPictureBox 
      Caption         =   "Picture Box control:"
      Height          =   195
      Left            =   3360
      TabIndex        =   21
      Top             =   2580
      Width           =   3015
   End
   Begin VB.Label lblTextBox 
      Caption         =   "Text Box control:"
      Height          =   195
      Left            =   3360
      TabIndex        =   19
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Caption         =   "Flat Combo box:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Caption         =   "Flat Dropdown combo:"
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Caption         =   "Flat Combo box (larger font):"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmTest.frx":2064
      Height          =   615
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   660
      Width           =   6315
   End
   Begin VB.Label lblSplash 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   6375
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private m_cFlatten() As cFlatControl
Private m_iCount As Long

Private Sub chkEnabled_Click()
Dim i As Long
Dim bEnabled As Boolean
Dim ctl As Control
   For Each ctl In Me.Controls
      On Error Resume Next
      If TypeName(ctl) <> "CheckBox" Then
         bEnabled = (chkEnabled.Value = Checked)
         ctl.Enabled = bEnabled
         If (TypeOf ctl Is ComboBox) Or (TypeOf ctl Is TextBox) Then
            ctl.BackColor = IIf(bEnabled, vbWindowBackground, vbButtonFace)
         End If
      End If
   Next ctl
End Sub

Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub Form_Load()
Dim ctl As Control
Dim bDoIt As Boolean
Dim i As Long

   For Each ctl In Me.Controls
      bDoIt = False
      If TypeOf ctl Is ComboBox Then
         bDoIt = True
      ElseIf TypeOf ctl Is TextBox Then
         ctl.Text = ctl.Name & ", vbAccelerator"
         bDoIt = True
      ElseIf TypeOf ctl Is PictureBox Then
         bDoIt = True
      End If
      If (bDoIt) Then
         m_iCount = m_iCount + 1
         ReDim Preserve m_cFlatten(1 To m_iCount) As cFlatControl
         Set m_cFlatten(m_iCount) = New cFlatControl
         m_cFlatten(m_iCount).Attach ctl
      End If
      If TypeOf ctl Is ComboBox Then
         For i = 1 To 20
            ctl.AddItem ctl.Name & ",Test Item " & i
         Next i
         ctl.ListIndex = 0
      End If
   Next ctl

End Sub

Private Sub optStyle_Click(Index As Integer)
Dim i As Long
   For i = 1 To m_iCount
      m_cFlatten(i).FlatStyle = Index
   Next i

End Sub

