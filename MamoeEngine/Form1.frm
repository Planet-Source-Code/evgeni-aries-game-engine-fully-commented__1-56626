VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Mamoe Engine"
   ClientHeight    =   9870
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOptions 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1560
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   551
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   703
      TabIndex        =   6
      Top             =   720
      Width           =   10575
      Begin VB.TextBox txtChat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   240
         Top             =   8400
         Width           =   9255
      End
   End
   Begin VB.PictureBox picTileBck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicRealBck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   9210
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   120
      Picture         =   "Form1.frx":007D
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Top             =   9120
      Width           =   10575
   End
   Begin VB.Label mnuOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   375
      Width           =   735
   End
   Begin VB.Label MenuBar 
      BackColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   15
      TabIndex        =   4
      Top             =   375
      Width           =   10770
   End
   Begin VB.Image imgMinBtnHover 
      Height          =   300
      Left            =   9975
      Picture         =   "Form1.frx":043E
      Top             =   45
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseBtnHover 
      Height          =   300
      Left            =   10350
      Picture         =   "Form1.frx":051D
      Top             =   45
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseBtn 
      Height          =   300
      Left            =   10350
      Picture         =   "Form1.frx":05FE
      Top             =   45
      Width           =   300
   End
   Begin VB.Image imgMinBtn 
      Height          =   300
      Left            =   9975
      Picture         =   "Form1.frx":06E3
      Top             =   45
      Width           =   300
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   165
      Left            =   9840
      Top             =   195
      Width           =   945
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   180
      Left            =   9840
      Top             =   15
      Width           =   945
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   9825
      Top             =   0
      Width           =   960
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   105
      Left            =   15
      Top             =   255
      Width           =   10770
   End
   Begin VB.Label formcaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mamoe Engine Prototype"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   9855
   End
   Begin VB.Shape Shape7 
      Height          =   9390
      Left            =   0
      Top             =   480
      Width           =   10800
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   135
      Left            =   15
      Top             =   120
      Width           =   10770
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   15
      Top             =   15
      Width           =   10770
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Keys As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub cmdSend_Click()
    MsgBox "This function hasn't been implented yet!"
    'modEngine.SendMessage (frmMain.txtMessage.Text)
End Sub
Private Sub Command2_Click()
    modEngine.FrameSpeed = Text2.Text
End Sub
Private Sub Form_Load()
    frmMain.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(frmMain.hWnd, &HA1, 2, 0&)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhoverGraphics
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AniRunning = False
    Unload Me
End Sub

Private Sub imgCloseBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.imgCloseBtnHover.Visible = True
End Sub
Private Sub imgCloseBtnHover_Click()
    Unload frmMain
End Sub
Private Sub imgMinBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.imgMinBtnHover.Visible = True
End Sub
Private Sub imgMinBtnHover_Click()
    frmMain.WindowState = vbMinimized
End Sub

Private Sub FormCaption_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(frmMain.hWnd, &HA1, 2, 0&)
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.imgCloseBtnHover.Visible = False
    frmMain.imgMinBtnHover.Visible = False
End Sub
Private Sub MenuBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhoverGraphics
End Sub

Private Sub mnuOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmOptions.Visible = True
End Sub
Private Sub OptPthJmp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OptPthJmp.ForeColor = vbBlue
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys = Keys & KeyCode & ";"
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call modEngine.CheckKey(Keys)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhoverGraphics
End Sub
Sub UnhoverGraphics()
    frmMain.imgCloseBtnHover.Visible = False
    frmMain.imgMinBtnHover.Visible = False
    frmMain.frmOptions.Visible = False
End Sub
Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then '// if enter
        '// call sub that will send the message
        MsgBox "This function has no been implented yet!", vbExclamation
    End If
End Sub
