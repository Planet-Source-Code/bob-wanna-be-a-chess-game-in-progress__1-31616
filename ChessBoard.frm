VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CJCChess"
   ClientHeight    =   7755
   ClientLeft      =   2460
   ClientTop       =   540
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6405
   Begin VB.PictureBox Picture1 
      Height          =   5730
      Left            =   4845
      Picture         =   "ChessBoard.frx":0000
      ScaleHeight     =   5670
      ScaleWidth      =   1485
      TabIndex        =   0
      Top             =   0
      Width           =   1545
      Begin VB.CommandButton Command10 
         Caption         =   "Move"
         Height          =   885
         Left            =   0
         Picture         =   "ChessBoard.frx":79F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   540
         TabIndex        =   1
         Top             =   2625
         Width           =   390
      End
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   31
      Left            =   750
      TabIndex        =   34
      Top             =   2040
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   31
      Left            =   435
      Top             =   5790
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   30
      Left            =   270
      Top             =   5970
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   29
      Left            =   1095
      Top             =   6150
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   28
      Left            =   1110
      Top             =   5895
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   27
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   26
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   25
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   24
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   23
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   22
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   21
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   20
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   19
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   18
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   17
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   16
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   15
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   14
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   13
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   12
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   11
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   10
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   9
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   8
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   7
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   6
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   5
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   4
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   3
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   2
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   1
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      Height          =   600
      Index           =   0
      Left            =   30
      Top             =   5790
      Visible         =   0   'False
      Width           =   505
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   4815
      X2              =   -45
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Line Line3 
      X1              =   4785
      X2              =   -45
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   4815
      X2              =   4815
      Y1              =   15
      Y2              =   5715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000007&
      X1              =   4800
      X2              =   4800
      Y1              =   15
      Y2              =   5715
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   700
      Left            =   2700
      Top             =   2205
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   30
      Left            =   750
      TabIndex        =   33
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   29
      Left            =   750
      TabIndex        =   32
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   28
      Left            =   750
      TabIndex        =   31
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   27
      Left            =   750
      TabIndex        =   30
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   26
      Left            =   750
      TabIndex        =   29
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   25
      Left            =   750
      TabIndex        =   28
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   24
      Left            =   750
      TabIndex        =   27
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   23
      Left            =   750
      TabIndex        =   26
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   22
      Left            =   750
      TabIndex        =   25
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   21
      Left            =   750
      TabIndex        =   24
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   20
      Left            =   750
      TabIndex        =   23
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   19
      Left            =   750
      TabIndex        =   22
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   18
      Left            =   750
      TabIndex        =   21
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   17
      Left            =   750
      TabIndex        =   20
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   16
      Left            =   750
      TabIndex        =   19
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   15
      Left            =   750
      TabIndex        =   18
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   14
      Left            =   750
      TabIndex        =   17
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   13
      Left            =   750
      TabIndex        =   16
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   12
      Left            =   750
      TabIndex        =   15
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   11
      Left            =   750
      TabIndex        =   14
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   10
      Left            =   750
      TabIndex        =   13
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   9
      Left            =   750
      TabIndex        =   12
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   8
      Left            =   750
      TabIndex        =   11
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   7
      Left            =   750
      TabIndex        =   10
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   6
      Left            =   750
      TabIndex        =   9
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   5
      Left            =   750
      TabIndex        =   8
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   4
      Left            =   750
      TabIndex        =   7
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   3
      Left            =   750
      TabIndex        =   6
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   2
      Left            =   750
      TabIndex        =   5
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   1
      Left            =   750
      TabIndex        =   4
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Piece 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      King"
      Height          =   705
      Index           =   0
      Left            =   750
      TabIndex        =   3
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   31
      Left            =   4020
      Picture         =   "ChessBoard.frx":7E36
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   30
      Left            =   4020
      Picture         =   "ChessBoard.frx":C444
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   29
      Left            =   4020
      Picture         =   "ChessBoard.frx":10A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   28
      Left            =   4020
      Picture         =   "ChessBoard.frx":15060
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   27
      Left            =   4020
      Picture         =   "ChessBoard.frx":1966E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   26
      Left            =   4020
      Picture         =   "ChessBoard.frx":1DC7C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   25
      Left            =   4020
      Picture         =   "ChessBoard.frx":2228A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   24
      Left            =   4020
      Picture         =   "ChessBoard.frx":26898
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   23
      Left            =   4020
      Picture         =   "ChessBoard.frx":2AEA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   22
      Left            =   4020
      Picture         =   "ChessBoard.frx":2F4B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   21
      Left            =   4020
      Picture         =   "ChessBoard.frx":33AC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   20
      Left            =   4020
      Picture         =   "ChessBoard.frx":380D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   19
      Left            =   4020
      Picture         =   "ChessBoard.frx":3C6DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   18
      Left            =   4020
      Picture         =   "ChessBoard.frx":40CEC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   17
      Left            =   4020
      Picture         =   "ChessBoard.frx":452FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   16
      Left            =   4020
      Picture         =   "ChessBoard.frx":49908
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   15
      Left            =   4020
      Picture         =   "ChessBoard.frx":4DF16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   14
      Left            =   4020
      Picture         =   "ChessBoard.frx":52524
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   13
      Left            =   4020
      Picture         =   "ChessBoard.frx":56B32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   12
      Left            =   4020
      Picture         =   "ChessBoard.frx":5B140
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   11
      Left            =   4020
      Picture         =   "ChessBoard.frx":5F74E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   10
      Left            =   4020
      Picture         =   "ChessBoard.frx":63D5C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   9
      Left            =   4020
      Picture         =   "ChessBoard.frx":6836A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   8
      Left            =   4020
      Picture         =   "ChessBoard.frx":6C978
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   7
      Left            =   4020
      Picture         =   "ChessBoard.frx":70F86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   6
      Left            =   4020
      Picture         =   "ChessBoard.frx":75594
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   5
      Left            =   4020
      Picture         =   "ChessBoard.frx":79BA2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   4
      Left            =   4020
      Picture         =   "ChessBoard.frx":7E1B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   3
      Left            =   4020
      Picture         =   "ChessBoard.frx":827BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   2
      Left            =   4020
      Picture         =   "ChessBoard.frx":86DCC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   1
      Left            =   4020
      Picture         =   "ChessBoard.frx":8B3DA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image WhiteS 
      Height          =   705
      Index           =   0
      Left            =   4020
      Picture         =   "ChessBoard.frx":8F9E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   31
      Left            =   2070
      Picture         =   "ChessBoard.frx":93FF6
      Stretch         =   -1  'True
      Top             =   315
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   30
      Left            =   2085
      Picture         =   "ChessBoard.frx":98278
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   29
      Left            =   2085
      Picture         =   "ChessBoard.frx":9C4FA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   28
      Left            =   2085
      Picture         =   "ChessBoard.frx":A077C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   27
      Left            =   2085
      Picture         =   "ChessBoard.frx":A49FE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   26
      Left            =   2085
      Picture         =   "ChessBoard.frx":A8C80
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   25
      Left            =   2085
      Picture         =   "ChessBoard.frx":ACF02
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   24
      Left            =   2085
      Picture         =   "ChessBoard.frx":B1184
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   23
      Left            =   2085
      Picture         =   "ChessBoard.frx":B5406
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   22
      Left            =   2085
      Picture         =   "ChessBoard.frx":B9688
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   21
      Left            =   2085
      Picture         =   "ChessBoard.frx":BD90A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   20
      Left            =   2085
      Picture         =   "ChessBoard.frx":C1B8C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   19
      Left            =   2085
      Picture         =   "ChessBoard.frx":C5E0E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   18
      Left            =   2085
      Picture         =   "ChessBoard.frx":CA090
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   17
      Left            =   2085
      Picture         =   "ChessBoard.frx":CE312
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   16
      Left            =   2085
      Picture         =   "ChessBoard.frx":D2594
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   15
      Left            =   2085
      Picture         =   "ChessBoard.frx":D6816
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   14
      Left            =   2085
      Picture         =   "ChessBoard.frx":DAA98
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   13
      Left            =   2085
      Picture         =   "ChessBoard.frx":DED1A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   12
      Left            =   2085
      Picture         =   "ChessBoard.frx":E2F9C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   11
      Left            =   2085
      Picture         =   "ChessBoard.frx":E721E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   10
      Left            =   2085
      Picture         =   "ChessBoard.frx":EB4A0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   9
      Left            =   2085
      Picture         =   "ChessBoard.frx":EF722
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   8
      Left            =   2085
      Picture         =   "ChessBoard.frx":F39A4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   7
      Left            =   2085
      Picture         =   "ChessBoard.frx":F7C26
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   6
      Left            =   2085
      Picture         =   "ChessBoard.frx":FBEA8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   5
      Left            =   2085
      Picture         =   "ChessBoard.frx":10012A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   4
      Left            =   2085
      Picture         =   "ChessBoard.frx":1043AC
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   3
      Left            =   2085
      Picture         =   "ChessBoard.frx":10862E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   2
      Left            =   2085
      Picture         =   "ChessBoard.frx":10C8B0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   1
      Left            =   2085
      Picture         =   "ChessBoard.frx":110B32
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Image BlackSquare 
      Height          =   705
      Index           =   0
      Left            =   2085
      Picture         =   "ChessBoard.frx":114DB4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BlackSquare_Click(Index As Integer)
DestX = BlackSquare(Index).Left / 600
DestY = BlackSquare(Index).Top / 715
Shape1.Left = DestX * 600
Shape1.Top = DestY * 715
Shape1.FillStyle = 1
Shape1.Visible = True
If O <> -1 Then
If P(O).Square(1) <> -1 Then
Command10_Click
End If
End If
End Sub

Private Sub Form_Load()
For i = 0 To 31
Taken(i) = -1
Next
ResetP
i = 0
Module1.Turn = 0
NS = 0
ResetTiles
For b = 0 To 31
P(b).Square(1) = -1
P(b).Square(2) = -1
Piece(b).Visible = True
Next
P(32).Square(1) = -1
P(32).Square(2) = -1
For b = 0 To 8
For i = 0 To 8
Board.Squares(i, b) = -1
Next i
Next b
SetPosAll
Turn = 0
LastBM = -1
LastWM = -1
Me.Show
DisableNotTurn
Piece_Click WhitePawn1 + 3
KH = Array(2, 1, -1, -2, -2, -1, 1, 2)
KV = Array(-1, -2, -2, -1, 1, 2, 2, 1)
End Sub

Public Sub SetPos(Obj, X, Y, Optional CheckWin = False)
On Error Resume Next
NX = P(Obj).Square(1)
NY = P(Obj).Square(2)
Turn = Not Turn
Board.Squares(NX, NY) = -1 ' Empty old pos
P(Obj).Square(1) = X
P(Obj).Square(2) = Y
' Avoid flickering from size change every time
' A piece is moved
If Piece(Obj).Width <> 600 Then Piece(Obj).Width = 600
If Piece(Obj).Height <> 700 Then Piece(Obj).Height = 700
Obj1 = Board.Squares(X, Y)
If CheckWin = True And Obj1 > -1 Then
   P(Obj1).Square(1) = -1
   P(Obj1).NExist = True
   If LastBM = Obj1 Then LastBM = -1
   If LastWM = Obj1 Then LastWM = -1
   If P(Obj1).T = King Then
      If Obj1 < 16 And Err.Number < 1 Then
         MsgBox "Black Wins!"
         Form_Load
         Exit Sub
      ElseIf Obj1 > 15 Then
         MsgBox "White Wins!"
         Form_Load
         Exit Sub
      End If
   End If
End If
P(Obj1).Square(1) = -1
P(Obj1).Square(2) = -1
'Piece(Obj1).Visible = False
Piece(Obj1).Left = X * 600
Piece(Obj1).Top = Y * 715

'Board.Squares(X, Y) = -1
Board.Squares(X, Y) = Obj
Piece(Obj).Left = X * 600
Piece(Obj).Top = Y * 715
If P(Obj).T = Pawn And (Y = 0 Or Y = 7) Then
2
   Select Case LCase(InputBox("Please enter " & _
   "Queen, Rook, Knight or Bishop: "))
      Case "queen"
         P(Obj).T = Queen
         Piece(Obj).Caption = vbCrLf & "Queen"
      Case "rook"
         P(Obj).T = Rook
         Piece(Obj).Caption = vbCrLf & "Rook"
      Case "bishop"
         P(Obj).T = Bishop
         Piece(Obj).Caption = vbCrLf & "Bishop"
      Case "knight"
         P(Obj).T = Knight
         Piece(Obj).Caption = vbCrLf & "Knight"
      Case Else
         GoTo 2
   End Select
End If
DisableNotTurn
UpdateTaken
End Sub

Function CanMove(Obj, X, Y)
CanMove = False
If X < 0 Or Y < 0 Or X >= 8 Or Y >= 8 Then
   CanMove = False
   GoTo 1
End If
Select Case P(Obj).T
   Case King
      If DiagCheck(Obj, X, Y, 1, "All") = True _
      Or LC(Obj, X, Y, 1, "King") = True Then
         CanMove = True
         GoTo 1
      End If
   Case Queen
      A = DiagCheck(Obj, DestX, DestY, 16, "All")
      If A = False Then A = LC(Obj, DestX, DestY, 16, "All")
      If A = True Then CanMove = True
   Case Bishop
      A = DiagCheck(Obj, DestX, DestY, 16, "All")
      If A = True Then CanMove = True
   Case Rook
      A = LC(Obj, DestX, DestY, 16, "All")
      If A = True Then CanMove = True
   Case Pawn
      A = LC(Obj, X, Y, 1, "Pawn")
      If A = True Then CanMove = True
   Case Knight
      For i = 0 To 7
         If DestY = P(O).Square(2) + KV(i) Then
            If DestX = P(O).Square(1) + KH(i) Then
               A = KnightCheck(Obj, i)
               If A = True Then CanMove = True
            End If
         End If
      Next
End Select
1
End Function

Function DiagCheck(Obj, X, Y, MaxMove, Dirs)
On Error Resume Next
Err.Clear
X1 = P(Obj).Square(1)
Y1 = P(Obj).Square(2)
If X1 = -1 Or Y1 = -1 Then GoTo 2
Direc = ""
If Y1 < Y Then
   Direc = Direc & "S"
   HX = Y
   LX = Y1
ElseIf Y1 > Y Then
   Direc = Direc & "N"
   HX = Y1
   LX = Y
End If
If X1 < X Then
   Direc = Direc & "E"
   HX = X
   LX = X1
ElseIf X1 > X Then
   Direc = Direc & "W"
   HX = X1
   LX = X
End If
If Dirs = "All" Then
MM = HX - LX + 1
If MM > MaxMove Then MM = MaxMove
   For i = 1 To MM
      If Direc = "NW" Then
         Obj1 = Board.Squares(X1 - i, Y1 - i)
         If CanGo(Obj, Obj1, MM, i) = False Then GoTo 1
      ElseIf Direc = "SE" Then
         Obj1 = Board.Squares(X1 + i, Y1 + i)
         If CanGo(Obj, Obj1, MM, i) = False Then GoTo 1
      ElseIf Direc = "NE" Then
         Obj1 = Board.Squares(X1 + i, Y1 - i)
         If CanGo(Obj, Obj1, MM, i) = False Then GoTo 1
      ElseIf Direc = "SW" Then
         Obj1 = Board.Squares(X1 - i, Y1 + i)
         If CanGo(Obj, Obj1, MM, i) = False Then GoTo 1
      End If
      If i > 1 And i >= MaxMove Then GoTo 2
      If (X = X1 - i And Y = Y1 - i) _
      Or (X = X1 + i And Y = Y1 + i) _
      Or (X = X1 + i And Y = Y1 - i) _
      Or (X = X1 - i And Y = Y1 + i) Then
         DiagCheck = True
         GoTo 2
      End If
   Next
End If
2
Exit Function
DiagCheck = False
1
End Function
Function LC(Obj, X, Y, Moves, Dirs)
Dim Done As Boolean
LC = False
X1 = P(Obj).Square(1)
Y1 = P(Obj).Square(2)
HY = Y
LY = Y
HX = X
LX = X
If Y1 > Y Then
   HY = Y1
   LY = Y
End If
If Y1 < Y Then
   HY = Y
   LY = Y1
End If
If X1 > X Then
   HX = X1
   LX = X
End If
If X1 < X Then
   HX = X
   LX = X1
End If
b = X1
If LY = -1 Or LX = -1 Then
   LC = False
   Exit Function
End If
For b = LY To HY
   For i = LX To HX
      If CanGoL(Obj, Board.Squares(i, b), i, b, X, Y) = False Then
         LC = False
         GoTo 1
      End If
   Next
Next

If Dirs = "All" Then
If (X > X1 And Y = Y1) _
Or (X < X1 And Y = Y1) _
Or (X = X1 And Y > Y1) _
Or (X = X1 And Y < Y1) Then
   LC = True
End If
ElseIf Dirs = "Pawn" Then
   If (X = X1 And Y = Y1 - 1 And O < 10 And O > 0) Then LC = True
   If (X = X1 And Y = Y1 + 1 And O > 10 And O < 32) Then LC = True
   If (X = X1 And Y = Y1 + 2 And P(O).Square(2) = 1) Then LC = True
   If (X = X1 And Y = Y1 - 2 And P(O).Square(2) = 6) Then LC = True
ElseIf Dirs = "King" Then
   If (X = X1 - 1 And Y = Y1 - 1) Then LC = True
   If (X = X1 + 1 And Y = Y1 + 1) Then LC = True
   If (X = X1 + 1 And Y = Y1 - 1) Then LC = True
   If (X = X1 - 1 And Y = Y1 + 1) Then LC = True
   If (X = X1 And Y = Y1 + 1) Then LC = True
   If (X = X1 And Y = Y1 - 1) Then LC = True
   If (X = X1 + 1 And Y = Y1) Then LC = True
   If (X = X1 - 1 And Y = Y1) Then LC = True
End If
1
End Function

Public Sub Piece_Click(Index As Integer)
For b = 0 To 31
   If Taken(b) = Index Then Exit Sub
Next
If (O <> -1) Then
   If (O < 16) Then Piece(O).ForeColor = vbBlack
   If (O > 15) Then Piece(O).ForeColor = vbWhite
   If (O < 16) Then Piece(O).BackColor = vbWhite
   If (O > 15) Then Piece(O).BackColor = vbBlack
End If
O = Index
Piece(Index).ForeColor = vbYellow
Piece(Index).BackColor = &H808080
End Sub
Sub ResetTiles()
Dim OffSet As Boolean
For R = 0 To 7
   For X = 0 To 3
      BlackSquare(NS).Left = (X * 2) * 600
      BlackSquare(NS).Top = i * 715
      If OffSet Then
      BlackSquare(NS).Left = BlackSquare(NS).Left + 600
      If X = 4 Then GoTo 1
      If X = 3 And R = 7 Then GoTo 1
      End If
      NS = NS + 1
   Next
1
   OffSet = Not OffSet
   i = i + 1
Next
OffSet = True
NS = 0
i = 0
For R = 0 To 7
   For X = 0 To 3
      WhiteS(NS).Left = (X * 2) * 600
      WhiteS(NS).Top = i * 715
      If OffSet Then
         WhiteS(NS).Left = WhiteS(NS).Left + 600
         If X = 4 Then GoTo 2
         If X = 3 And R = 7 Then GoTo 1
      End If
      NS = NS + 1
   Next
2
   OffSet = Not OffSet
   i = i + 1
Next

End Sub

Private Sub WhiteS_Click(Index As Integer)
DestX = WhiteS(Index).Left / 600
DestY = WhiteS(Index).Top / 715
Shape1.Left = DestX * 600
Shape1.Top = DestY * 715
Shape1.FillStyle = 1
Shape1.Visible = True
If O <> -1 Then
   If P(O).Square(1) <> -1 Then
      Command10_Click
   End If
End If

End Sub

Function KnightCheck(Obj, MN)
On Error Resume Next
KnightCheck = True
If DestX > 8 Or DestX < 0 Or _
DestY > 8 Or DestY < 0 Then
   KnightCheck = False
   If CanGo(Obj, Board.Squares(DestX, DestY), 1, 1) = False Then _
   KnightCheck = False
End If
End Function
Private Sub Command10_Click()
If O = -1 Then Exit Sub
M = 1
PosX = P(O).Square(1) - M
PosY = P(O).Square(2) - M
If CanMove(O, DestX, DestY) Then
   SetPos O, DestX, DestY, True
   Shape1.Visible = False
   Shape1.FillStyle = 1
   Label1 = ""
Else
   Label1 = "Illegal Move."
   Shape1.FillStyle = 7
End If
End Sub

Function CanGo(Obj, Obj1, Optional MM = -1, Optional i = -1)
CanGo = True
If Obj >= 16 Then Team = 1
If Obj <= 15 Then Team = 0
If Obj1 >= 16 Then Team2 = 1
If Obj1 <= 15 Then Team2 = 0
If Team = Team2 Then CanGo = False
If MaxMove <> -1 Then _
If i < MM - 1 Then CanGo = False
If i >= MM - 1 And Team <> Team2 Then CanGo = True
If Obj1 = -1 Then CanGo = True
If Obj1 = Obj Then CanGo = True
End Function
Function CanGoL(Obj, Obj1, X, Y, DestX, DestY)
CanGoL = True
If Obj >= 16 Then Team = 1
If Obj <= 15 Then Team = 0
If Obj1 >= 16 Then Team2 = 1
If Obj1 <= 15 Then Team2 = 0
If Team = Team2 Then CanGoL = False
If Obj1 <> Obj Then CanGoL = False
If MaxMove <> -1 Then _
If i < MM Then CanGoL = False
If X = DestX And Y = DestY And Team <> Team2 Then CanGoL = True
If Obj1 = -1 Then CanGoL = True
If Obj1 = Obj Then CanGoL = True
End Function

Sub UpdateTaken()
For i = 0 To 31
   If P(i).NExist = True Then
      If IB(i) Then
      Min = 16
      Max = 31
      Else
      Min = 0
      Max = 15
      End If
      For b = Min To Max
         If Taken(b) = i Or Taken(b) = -1 Then
            If IW(i) Then Piece(i).BackColor = vbWhite
            If IB(i) Then Piece(i).BackColor = vbBlack
            Piece(i).Left = Shape2(b).Left
            Piece(i).Top = Shape2(b).Top
            Piece(i).Width = 480
            Piece(i).Height = 580
            P(i).NExist = False
            Taken(b) = i
            GoTo 1
         End If
      Next
1
   End If
Next
End Sub
