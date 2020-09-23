VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mamoe Engine [Settings]"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame7 
         Caption         =   "Items and Furniture Codes"
         Height          =   1935
         Left            =   3720
         TabIndex        =   24
         Top             =   3120
         Width           =   2535
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Code = D"
            Height          =   255
            Left            =   1320
            TabIndex        =   26
            Top             =   1580
            Width           =   1095
         End
         Begin VB.Shape Shape4 
            Height          =   255
            Left            =   1320
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Shape Shape3 
            Height          =   1095
            Left            =   1320
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image Image2 
            Height          =   885
            Left            =   1440
            Picture         =   "frmSettings.frx":0000
            Stretch         =   -1  'True
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Code = I"
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
            Left            =   120
            TabIndex        =   25
            Top             =   1580
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            Height          =   255
            Left            =   120
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            Height          =   1095
            Left            =   120
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   840
            Left            =   240
            Picture         =   "frmSettings.frx":0823
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Character"
         Height          =   1935
         Left            =   3720
         TabIndex        =   22
         Top             =   1080
         Width           =   2535
         Begin VB.PictureBox Picture2 
            Height          =   1575
            Left            =   120
            Picture         =   "frmSettings.frx":096A
            ScaleHeight     =   1515
            ScaleWidth      =   555
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "WallPaper"
         Height          =   855
         Left            =   3720
         TabIndex        =   19
         Top             =   120
         Width           =   2535
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Text            =   "26,26,26"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "WallPaper Color RGB()"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Account"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   4800
         Width           =   3495
         Begin VB.TextBox txtAccount 
            Height          =   285
            Left            =   720
            MaxLength       =   12
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label5 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   260
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "HeightMap Values"
         Height          =   4695
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   2175
         Begin RichTextLib.RichTextBox MapRTF 
            Height          =   1335
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2355
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmSettings.frx":0D18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox CharRTF 
            Height          =   1095
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1931
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmSettings.frx":0D93
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox FurniRTF 
            Height          =   1215
            Left            =   120
            TabIndex        =   15
            Top             =   3360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2143
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmSettings.frx":0E0E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Furniture Setting"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Char Setting"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Map Setting"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tile Selection"
         Height          =   4695
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1335
         Begin VB.OptionButton optTile 
            Caption         =   "Gray White"
            Height          =   735
            Index           =   5
            Left            =   120
            Picture         =   "frmSettings.frx":0E89
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3840
            Width           =   1095
         End
         Begin VB.OptionButton optTile 
            Caption         =   "Green Tile"
            Height          =   735
            Index           =   4
            Left            =   120
            Picture         =   "frmSettings.frx":0FA3
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3120
            Width           =   1095
         End
         Begin VB.OptionButton optTile 
            Caption         =   "YellowTile"
            Height          =   735
            Index           =   3
            Left            =   120
            Picture         =   "frmSettings.frx":1084
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2400
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTile 
            Caption         =   "BlueTile"
            Height          =   735
            Index           =   2
            Left            =   120
            Picture         =   "frmSettings.frx":1165
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton optTile 
            Caption         =   "Red Tile"
            Height          =   735
            Index           =   1
            Left            =   120
            Picture         =   "frmSettings.frx":1250
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optTile 
            Caption         =   "Gray Tile"
            Height          =   735
            Index           =   0
            Left            =   120
            Picture         =   "frmSettings.frx":133A
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   5160
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R As Integer
Dim G As Integer
Dim B As Integer

Private Sub Command1_Click()
Dim MData() As String
Dim CData() As String
Dim FData() As String
Dim UboundData As Variant
    modEngine.AccountName = txtAccount.Text
    modEngine.AccountName = frmSettings.txtAccount.Text
    R = Split(frmSettings.Text1.Text, ",")(0)
    G = Split(frmSettings.Text1.Text, ",")(1)
    B = Split(frmSettings.Text1.Text, ",")(2)
    UboundData = Split(MapRTF.Text, vbNewLine)
    ReDim MData(1 To UBound(UboundData)) As String
    ReDim CData(1 To UBound(UboundData)) As String
    ReDim FData(1 To UBound(UboundData)) As String
    For X = 1 To UBound(UboundData)
        MData(X) = Split(MapRTF.Text, vbNewLine)(X - 1)
        CData(X) = Split(CharRTF.Text, vbNewLine)(X - 1)
        FData(X) = Split(FurniRTF.Text, vbNewLine)(X - 1)
    Next X
    WallColor = RGB(Int(R), Int(G), Int(B))
    Call modEngine.InitData(MData(), CData(), FData())
    If GameError = False Then
        Call modEngine.CreateRoom
    End If
End Sub

Private Sub Form_Load()
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOOO" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOXX" & vbNewLine
    MapRTF.Text = MapRTF.Text & "OOOOOOOOOXX" & vbNewLine
    
    CharRTF.Text = CharRTF.Text & "OXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    CharRTF.Text = CharRTF.Text & "XXXXXXXXXXX" & vbNewLine
    '//minibar
    'FurniRTF.Text = FurniRTF.Text & "00000IIIIII" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "0I000I0000I" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "0I000I0000I" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "0I000I0000I" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "0I000ID000I" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "0IIIIIIII0I" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "00000000000" & vbNewLine
    'FurniRTF.Text = FurniRTF.Text & "00000000000" & vbNewLine
    '//club
    '//yellow tile
    '//26,26,26 wallpaper
     FurniRTF.Text = FurniRTF.Text & "0IIIIIIIIII" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "00000ID000I" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "I0000ID000I" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "I0000I0000I" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "IIII0III0II" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "ID00000000I" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "ID00000000I" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "ID000000III" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "ID000000I00" & vbNewLine
     FurniRTF.Text = FurniRTF.Text & "IIIIIIIII00" & vbNewLine
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub optTile_Click(Index As Integer)
    Memory.picTile1.Picture = optTile(Index).Picture
End Sub

