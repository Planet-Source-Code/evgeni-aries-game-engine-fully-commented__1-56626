VERSION 5.00
Begin VB.Form Memory 
   Caption         =   "Form2"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBodyUpDrinkMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   4320
      Picture         =   "Memory.frx":0000
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   22
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox picBodyUpDrink 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   4320
      Picture         =   "Memory.frx":0194
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   21
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox picBodyLeftDrinkMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   3600
      Picture         =   "Memory.frx":0528
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox picBodyLeftDrink 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   3600
      Picture         =   "Memory.frx":0652
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox picMountinDewOpen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1665
      Left            =   960
      Picture         =   "Memory.frx":09CD
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   18
      Top             =   3480
      Width           =   1035
   End
   Begin VB.PictureBox picBodyDrinkRightMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3000
      Picture         =   "Memory.frx":1277
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   17
      Top             =   960
      Width           =   585
   End
   Begin VB.PictureBox picBodyDrinkRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3000
      Picture         =   "Memory.frx":137F
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   16
      Top             =   0
      Width           =   585
   End
   Begin VB.PictureBox picBodyDrinkDownMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   2400
      Picture         =   "Memory.frx":192F
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   15
      Top             =   960
      Width           =   585
   End
   Begin VB.PictureBox picBodyDrinkDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   2400
      Picture         =   "Memory.frx":1A36
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   14
      Top             =   0
      Width           =   585
   End
   Begin VB.PictureBox picItemBoxMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   900
      Left            =   1080
      Picture         =   "Memory.frx":1DF8
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   13
      Top             =   2520
      Width           =   870
   End
   Begin VB.PictureBox picItemBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   900
      Left            =   1080
      Picture         =   "Memory.frx":1EA9
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   12
      Top             =   2040
      Width           =   870
   End
   Begin VB.PictureBox picMountinDewMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1665
      Left            =   0
      Picture         =   "Memory.frx":1FF0
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   11
      Top             =   2640
      Width           =   1035
   End
   Begin VB.PictureBox picMountinDew 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1665
      Left            =   0
      Picture         =   "Memory.frx":2151
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   10
      Top             =   2040
      Width           =   1035
   End
   Begin VB.PictureBox PicTile1Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   0
      Picture         =   "Memory.frx":2937
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   9
      Top             =   4320
      Width           =   870
   End
   Begin VB.PictureBox picTile1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   0
      Picture         =   "Memory.frx":29CE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   8
      Top             =   4560
      Width           =   870
   End
   Begin VB.PictureBox picBodyUpMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   1800
      Picture         =   "Memory.frx":2AAF
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox picBodyUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   1800
      Picture         =   "Memory.frx":2BBC
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picBodyLeftMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   1200
      Picture         =   "Memory.frx":3039
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox picBodyLeft 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   1200
      Picture         =   "Memory.frx":3143
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picBodyRightMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "Memory.frx":3485
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   3
      Top             =   480
      Width           =   585
   End
   Begin VB.PictureBox picBodyRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "Memory.frx":358A
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   2
      Top             =   0
      Width           =   585
   End
   Begin VB.PictureBox picBodyDownMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   0
      Picture         =   "Memory.frx":3A34
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   1
      Top             =   480
      Width           =   585
   End
   Begin VB.PictureBox picBodyDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   0
      Picture         =   "Memory.frx":3B3B
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
