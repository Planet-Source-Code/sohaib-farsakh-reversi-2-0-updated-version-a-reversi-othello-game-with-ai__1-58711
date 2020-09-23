VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversi"
   ClientHeight    =   5805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7320
   ForeColor       =   &H00B75820&
   Icon            =   "Reversi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":3FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":40DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":430E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":4426
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":527A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":60CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reversi.frx":6F22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7335
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "Start a new game"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Load position"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save position"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo move"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo move"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "stop"
               Object.ToolTipText     =   "Stop auto play"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "show"
               Object.ToolTipText     =   "Show possible moves"
               ImageIndex      =   7
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sound"
               Object.ToolTipText     =   "Sound on"
               ImageIndex      =   8
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               Object.ToolTipText     =   "Help"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   360
      Top             =   5280
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Setup"
      Height          =   4815
      Left            =   5280
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1095
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1575
         Begin VB.OptionButton Option5 
            BackColor       =   &H0080C0FF&
            Caption         =   "Erase disk"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H0080C0FF&
            Caption         =   "White disc"
            Height          =   315
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H0080C0FF&
            Caption         =   "Black disc"
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "End Setup"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "White Player"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Black Player"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc type:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Side to play:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6240
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Status"
      Height          =   4815
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      Begin VB.TextBox txtMoves 
         Height          =   1935
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   840
         Top             =   3600
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Good Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Human"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "White Player:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Black Player:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Black discs:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "White discs:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Black"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Next turn:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   4440
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A           B           C           D           E           F           G           H"
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   600
      Width           =   4335
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   344
      X2              =   344
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   304
      X2              =   304
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   264
      X2              =   264
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   224
      X2              =   224
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   184
      X2              =   184
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   144
      X2              =   144
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   104
      X2              =   104
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   64
      X2              =   64
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   24
      X2              =   24
      Y1              =   56
      Y2              =   376
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   24
      X2              =   344
      Y1              =   376
      Y2              =   376
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   7
      X1              =   24
      X2              =   344
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   6
      X1              =   24
      X2              =   344
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   5
      X1              =   24
      X2              =   344
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   4
      X1              =   24
      X2              =   344
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   24
      X2              =   344
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   24
      X2              =   344
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   24
      X2              =   344
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   24
      X2              =   344
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   63
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   62
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   61
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   60
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   59
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   58
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   57
      Left            =   960
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   56
      Left            =   360
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   55
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   54
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   53
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   52
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   51
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   50
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   49
      Left            =   960
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   48
      Left            =   360
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   47
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   46
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   45
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   44
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   43
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   42
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   41
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   40
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   39
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   38
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   37
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   36
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   35
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   34
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   33
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   32
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   31
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   30
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   29
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   28
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   27
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   26
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   25
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   24
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   23
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   22
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   21
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   20
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   19
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   18
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   17
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   16
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   15
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   14
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   13
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   12
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   11
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   10
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   9
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   8
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   7
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   6
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   5
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   3
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   2
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   1
      Left            =   960
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   360
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin ComctlLib.ImageList imgMain 
      Left            =   4560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":7036
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":8348
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":965A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":A96C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnugame 
      Caption         =   "Game"
      Begin VB.Menu mnunew 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuload 
         Caption         =   "Load Position"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save Position"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusetup 
         Caption         =   "Setup Position"
      End
      Begin VB.Menu seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuundo 
         Caption         =   "Undo Move"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuredo 
         Caption         =   "Redo Move"
         Shortcut        =   ^Y
      End
      Begin VB.Menu seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuseggest 
         Caption         =   "Suggest Move"
         Shortcut        =   ^M
      End
      Begin VB.Menu seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnublack 
         Caption         =   "Black Player"
         Begin VB.Menu blackoptions 
            Caption         =   "Human"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu blackoptions 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Beginner Computer"
            Index           =   2
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Novice Computer"
            Index           =   3
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Moderate Computer"
            Index           =   4
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Good Computer"
            Index           =   5
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Expert Computer"
            Index           =   6
         End
      End
      Begin VB.Menu mnuwhite 
         Caption         =   "White Player"
         Index           =   0
         Begin VB.Menu whiteoptions 
            Caption         =   "Human"
            Index           =   0
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Beginner Computer"
            Index           =   2
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Novice Computer"
            Index           =   3
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Moderate Computer"
            Index           =   4
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Good Computer"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Expert Computer"
            Index           =   6
         End
      End
      Begin VB.Menu seperate1 
         Caption         =   "-"
      End
      Begin VB.Menu mnushowpossible 
         Caption         =   "Show possible moves"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusound 
         Caption         =   "Sound On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnudelay 
         Caption         =   "Delay before computer play"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'ensure all variables are declared

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'stop for a while
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'for playing sounds


Dim Grid(1 To 8, 1 To 8) As Integer 'keeps the board information(0 = nothing , 1 = black , 2 = white)
Dim turn As String 'which player is going to play now
Dim WhiteCount As Integer, BlackCount As Integer 'number of discs for each player
Dim LegalMoves(40, 2) As Integer, LegalMovesNum
Dim NewDiscs(1 To 30, 1 To 2) As Integer, FlippedNum As Integer
Dim SelMove(1 To 2) As Integer 'the move that the computer thinks its the best (i for x coordinate,2 for y coordinate)

Dim Nodes As Long 'number of nodes
Dim SearchEnd As Boolean 'true if we are at the end of game and will calculate the exact final score
Dim startdepth As Integer 'the depth we are going to search to
Dim MidDepth As Integer, EndDepth As Integer 'the depth for the mid game and for the end game(depends of the computer level)
Dim Player1Type As Integer, Player2Type As Integer 'whether the player is a human or a computer at a certain level
Dim freezed As Boolean
Dim MinDelay As Integer
Dim GridHistory(1 To 8, 1 To 8, 0 To 128) As Integer, MoveList(0 To 128) As Integer, MovesNum As Integer, MovesNum1 As Integer
Dim IsMouseDown As Boolean
Dim PlayBeginner As Boolean

Private Sub DrawBoard()
'draws the board
Dim imgnum As Integer, theplayer As Integer, i As Integer, j As Integer
Select Case turn
Case "black"
theplayer = 1
Case "white"
theplayer = 2
End Select
For i = 1 To 8 'loop over in x direction
    For j = 1 To 8 'loop over cells in y direction
    Select Case Grid(i, j)
    Case 0
    imgnum = 3 'image number 3 in the imagelist
    Case 1
    imgnum = 2 'image number 2 in the imagelist
    Case 2
    imgnum = 1 'image number 1 in the imagelist
    End Select
    'show possible moves
    If IsValid(theplayer, i, j) = True And mnushowpossible.Checked = True And mnusetup.Checked = False Then imgnum = 4
    'draw the image
    Image1((j - 1) * 8 + i - 1).Picture = imgMain.ListImages.Item(imgnum).Picture
    Next
Next
End Sub
Private Sub MakeMove(player As Integer, X As Integer, Y As Integer)
Dim i As Integer, j As Integer, xx As Integer, Flipped As Integer
'changes the grid information after making a move and updates the number of discs for each player


Grid(X, Y) = player

'check for flippedd discs above the new placed disc
Flipped = 0
For i = Y - 1 To 1 Step -1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player Then
        For j = Y - 1 To i + 1 Step -1
            Grid(X, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check for flippedd discs below the new placed disc
For i = Y + 1 To 8 Step 1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player Then
        For j = Y + 1 To i - 1 Step 1
            Grid(X, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check for flippedd discs at the right of the new placed disc
For i = X + 1 To 8 Step 1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player Then
        For j = X + 1 To i - 1 Step 1
            Grid(j, Y) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next


'check for flippedd discs at the left of the new placed disc\
For i = X - 1 To 1 Step -1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player Then
        For j = X - 1 To i + 1 Step -1
            Grid(j, Y) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check up left
For i = Y - 1 To 1 Step -1
    xx = X - (Y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y - 1 To i + 1 Step -1
            xx = X - (Y - j)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check up right
For i = Y - 1 To 1 Step -1
    xx = X + (Y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y - 1 To i + 1 Step -1
            xx = X + (Y - j)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check down left
For i = Y + 1 To 8 Step 1
    xx = X - (i - Y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y + 1 To i - 1 Step 1
            xx = X - (j - Y)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check down right
For i = Y + 1 To 8 Step 1
    xx = X + (i - Y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y + 1 To i - 1 Step 1
            xx = X + (j - Y)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'update the number of discs for each player
Select Case player
Case 2
WhiteCount = WhiteCount + Flipped + 1
BlackCount = BlackCount - Flipped
Case 1
WhiteCount = WhiteCount - Flipped
BlackCount = BlackCount + Flipped + 1
End Select

End Sub

Public Sub PlaySound(sFilename As String)
sFilename = sFilename & ".wav"
Call sndPlaySound(App.Path & "\" & sFilename, &H1)
End Sub


Private Sub GetFlipped(player As Integer, X As Integer, Y As Integer)
Dim i As Integer, j As Integer, xx As Integer
'this sub if used in move ordering, if fills the (NewDiscs) array
'with the discs that will be flipped after
'making a move and if works almost the same as the (MakeMove) sub

FlippedNum = 1

NewDiscs(1, 1) = X
NewDiscs(1, 2) = Y

For i = Y - 1 To 1 Step -1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player Then
        For j = Y - 1 To i + 1 Step -1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = X
        NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = Y + 1 To 8 Step 1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player Then
        For j = Y + 1 To i - 1 Step 1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = X
        NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = X + 1 To 8 Step 1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player Then
        For j = X + 1 To i - 1 Step 1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = j
        NewDiscs(FlippedNum, 2) = Y
        Next
        Exit For
    End If
Next

For i = X - 1 To 1 Step -1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player Then
        For j = X - 1 To i + 1 Step -1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = j
        NewDiscs(FlippedNum, 2) = Y
        Next
        Exit For
    End If
Next

For i = Y - 1 To 1 Step -1
    xx = X - (Y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y - 1 To i + 1 Step -1
            xx = X - (Y - j)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = Y - 1 To 1 Step -1
    xx = X + (Y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y - 1 To i + 1 Step -1
            xx = X + (Y - j)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = Y + 1 To 8 Step 1
    xx = X - (i - Y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y + 1 To i - 1 Step 1
            xx = X - (j - Y)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = Y + 1 To 8 Step 1
    xx = X + (i - Y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = Y + 1 To i - 1 Step 1
            xx = X + (j - Y)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next


End Sub

Private Function IsValid(player As Integer, X As Integer, Y As Integer) As Boolean
Dim i As Integer, xx As Integer
'this function checkes whether a move if valid in the current position and for a certain player
'it is similar to the(MakeMove)sub but if doesn't care about how many discs are flipped, it breaks
'and returnes true at the first time if finds a disc that wil be flipped

IsValid = False
If Grid(X, Y) <> 0 Then Exit Function
For i = Y - 1 To 1 Step -1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player And (Y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(X, i) = player Then Exit For
Next

For i = Y + 1 To 8 Step 1
    If Grid(X, i) = 0 Then Exit For
    If Grid(X, i) = player And (i - Y) >= 2 Then IsValid = True: Exit Function
    If Grid(X, i) = player Then Exit For
Next

For i = X + 1 To 8 Step 1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player And (i - X) >= 2 Then IsValid = True: Exit Function
    If Grid(i, Y) = player Then Exit For
Next

For i = X - 1 To 1 Step -1
    If Grid(i, Y) = 0 Then Exit For
    If Grid(i, Y) = player And (X - i) >= 2 Then IsValid = True: Exit Function
    If Grid(i, Y) = player Then Exit For
Next

For i = Y - 1 To 1 Step -1
    xx = X - (Y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (Y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = Y - 1 To 1 Step -1
    xx = X + (Y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (Y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = Y + 1 To 8 Step 1
    xx = X - (i - Y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (i - Y) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = Y + 1 To 8 Step 1
    xx = X + (i - Y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (i - Y) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

End Function

Private Function CountBad(X As Integer, Y As Integer)
Dim Bad As Integer
'counts how many empty cells are around a cell,it is used for move ordering
'so the more empty discs around a cell, the worse it is
Bad = 0
If X > 1 Then
If Grid(X - 1, Y) = 0 Then Bad = Bad + 1
End If

If X < 8 Then
If Grid(X + 1, Y) = 0 Then Bad = Bad + 1
End If

If Y > 1 Then
If Grid(X, Y - 1) = 0 Then Bad = Bad + 1
End If

If Y < 8 Then
If Grid(X, Y + 1) = 0 Then Bad = Bad + 1
End If

If X > 1 And Y > 1 Then
If Grid(X - 1, Y - 1) = 0 Then Bad = Bad + 1
End If

If X < 8 And Y < 8 Then
If Grid(X + 1, Y + 1) = 0 Then Bad = Bad + 1
End If

If X > 1 And Y < 8 Then
If Grid(X - 1, Y + 1) = 0 Then Bad = Bad + 1
End If

If X < 8 And Y > 1 Then
If Grid(X + 1, Y - 1) = 0 Then Bad = Bad + 1
End If

CountBad = Bad

End Function

Private Sub GetLegalMoves(player As Integer)
'fills the (LegalMoves) array with the legal moves for a certain
'player in the current position(used for move generation)
Dim i As Integer, j As Integer
LegalMovesNum = 0
For i = 1 To 8
    For j = 1 To 8
    If IsValid(player, i, j) = True Then
    LegalMovesNum = LegalMovesNum + 1
    LegalMoves(LegalMovesNum, 1) = i
    LegalMoves(LegalMovesNum, 2) = j
    End If
    Next
Next
End Sub

Private Sub UpdateMoveList()
Dim i As Integer
txtMoves.Text = ""
For i = 1 To MyInt(MovesNum / 2)
txtMoves.SelStart = Len(txtMoves.Text)
If MovesNum Mod 2 = 0 Or i < MyInt(MovesNum / 2) Then
txtMoves.Text = txtMoves.Text & Str(i * 2 - 1) & "." & NumberToLetters(MoveList(i * 2 - 1)) & vbTab & Str(i * 2) & "." & NumberToLetters(MoveList(i * 2)) & vbCrLf
Else
txtMoves.Text = txtMoves.Text & Str(i * 2 - 1) & "." & NumberToLetters(MoveList(i * 2 - 1))
End If
Next
txtMoves.SelStart = Len(txtMoves.Text)
End Sub
Private Function MyInt(number As Single) As Integer
Dim a As Integer
a = Int(number)
If a = number Then MyInt = a Else MyInt = a + 1
End Function
Private Function NumberToLetters(number As Integer)
If number = -1 Then NumberToLetters = "----": Exit Function
Dim a As Integer, b As Integer
a = (number Mod 8) + 1
b = Int(number / 8) + 1
Dim Letter As String * 1
Select Case a
Case 1
Letter = "A"
Case 2
Letter = "B"
Case 3
Letter = "C"
Case 4
Letter = "D"
Case 5
Letter = "E"
Case 6
Letter = "F"
Case 7
Letter = "G"
Case 8
Letter = "H"
End Select
NumberToLetters = Letter & Trim(Str(b))

End Function
Private Sub NewGame()
'starts a new game, resets everything to default
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    Grid(i, j) = 0
    Next
Next
Grid(4, 4) = 2
Grid(5, 5) = 2
Grid(5, 4) = 1
Grid(4, 5) = 1
WhiteCount = 2
BlackCount = 2
Label5.Caption = "2"
Label6.Caption = "2"

turn = "black"
Label2.Caption = "black"
DrawBoard

Timer1.Enabled = True
Timer2.Enabled = True
freezed = False
IsMouseDown = False

MovesNum = 0
MovesNum1 = 0
For i = 1 To 8
For j = 1 To 8
GridHistory(i, j, 0) = Grid(i, j)
Next
Next
UpdateMoveList
If mnusound.Checked = True Then PlaySound ("NewGame")

End Sub
Private Function EvaluateEnd(player As Integer)
'the evaluation function used in the end of the game, it returnes
'the difference between our discs and the opponents discs
Dim Value As Integer
Value = WhiteCount - BlackCount
If player = 2 Then EvaluateEnd = Value Else: EvaluateEnd = -Value
End Function

Private Function Evaluate(player As Integer)
'the evaluation function used in the begining and the middle of the game, it is
'based on mobility(number of legal moves) for each player and also on the board
'positions that each player occupy
Dim Value As Integer, cellval As Integer, cellnum As Integer
Dim i As Integer, j As Integer, Counted As Integer
Dim HasMove1 As Boolean, HasMove2 As Boolean
HasMove1 = False: HasMove2 = False
'at first the function evaluates for the white player and then changes the sign if we are evaluating for black
Value = 0
For i = 1 To 8
    For j = 1 To 8
    If IsValid(2, i, j) = True Then
        Value = Value + 18: HasMove2 = True 'add 18 for each legal move for white
        If (i = 1 And j = 1) Or (i = 1 And j = 8) Or (i = 8 And j = 1) Or (i = 8 And j = 8) Then
            If player = 2 Then Value = Value + 150 Else Value = Value + 10
        End If
    End If
    
    If IsValid(1, i, j) = True Then
        Value = Value - 18: HasMove1 = True 'subtract 18 for each legal move for black
        If (i = 1 And j = 1) Or (i = 1 And j = 8) Or (i = 8 And j = 1) Or (i = 8 And j = 8) Then
            If player = 1 Then Value = Value - 150 Else Value = Value - 10
        End If
    End If
    
    If Grid(i, j) <> 0 Then
        cellnum = (j - 1) * 8 + i
        
        cellval = 0
        Select Case cellnum
        Case 1, 8, 57, 64 'a corner
        cellval = 250
        Case 2, 7, 9, 16, 49, 56, 58, 63 'C square(a square that is adjacent to a corner horizontally or vertically
            If cellnum = 2 Or cellnum = 9 Then
                If Grid(1, 1) = Grid(i, j) Then cellval = 15
                If Grid(1, 1) = 0 Then cellval = -30
            End If
            If cellnum = 7 Or cellnum = 16 Then
                If Grid(8, 1) = Grid(i, j) Then cellval = 15
                If Grid(8, 1) = 0 Then cellval = -30
            End If
            If cellnum = 49 Or cellnum = 58 Then
                If Grid(1, 8) = Grid(i, j) Then cellval = 15
                If Grid(1, 8) = 0 Then cellval = -30
            End If
            If cellnum = 56 Or cellnum = 63 Then
                If Grid(8, 8) = Grid(i, j) Then cellval = 15
                If Grid(8, 8) = 0 Then cellval = -30
            End If
            
        Case 10, 15, 50, 55 'X square(a square that is adjacent to a corner diagonally
            
            If cellnum = 10 Then
                If Grid(1, 1) = Grid(i, j) Then cellval = 5
                If Grid(1, 1) = 0 Then cellval = -120
            End If
            If cellnum = 15 Then
                If Grid(8, 1) = Grid(i, j) Then cellval = 5
                If Grid(8, 1) = 0 Then cellval = -120
            End If
            If cellnum = 50 Then
                If Grid(1, 8) = Grid(i, j) Then cellval = 5
                If Grid(1, 8) = 0 Then cellval = -120
            End If
            If cellnum = 55 Then
                If Grid(8, 8) = Grid(i, j) Then cellval = 5
                If Grid(8, 8) = 0 Then cellval = -120
            End If
            
        Case 3, 4, 5, 6, 17, 25, 33, 41, 24, 32, 40, 48, 59, 60, 61, 62 ' sides other than C squares
        cellval = 15
        Case Else 'any other place(in the middle of the board)
        cellval = 0
        Counted = CountBad(i, j)
        cellval = cellval - (Counted + 3 * Sgn(Counted)) * (64 - (WhiteCount + BlackCount)) / 20 ' the value of the cell also depends on how many empty cells are surrounding it
        End Select
        'adds or subtracts the score depending on which player occupies that cell
        If Grid(i, j) = 2 Then Value = Value + cellval Else: Value = Value - cellval
    End If
    Next
Next
If HasMove1 = False And HasMove2 = False Then
    If WhiteCount > BlackCount Then Value = 2500
    If WhiteCount < BlackCount Then Value = -2500
    If WhiteCount = BlackCount Then Value = 0
End If

'negate the value if we are evaluating for black and return the normal value if we are evaluating for white
If player = 2 Then Evaluate = Value Else: Evaluate = -Value
End Function


Private Function Search(depth As Integer, alpha As Integer, beta As Integer)
'the searching function(finds the best move for the computer)
Dim NewScore As Integer
Dim LegalMovesNow(40, 2) As Integer 'keeps the legal moves in this procedure
Dim player As Integer 'which player is going to play at the current depth
Dim TempGrid(1 To 8, 1 To 8) As Integer 'stores the grid information before making a move in order to be able to unmake it
Dim a As Integer, b As Integer
Dim tempwhite As Integer, tempblack As Integer 'stores the number of discs for each player before making a move in order to be able to unmake it
Dim i As Integer

'determine which player is to go at the current depth
Select Case turn
Case "white"
If (startdepth - depth) Mod 2 = 0 Then
player = 2
Else
player = 1
End If
Case "black"
If (startdepth - depth) Mod 2 = 0 Then
player = 1
Else
player = 2
End If
End Select

Erase LegalMovesNow()

If depth = 0 Then 'if we are at the end of the tree then return evaluation
    If SearchEnd = True Then
    alpha = EvaluateEnd(player)
    Else
    alpha = Evaluate(player)
    End If
    Search = alpha
    Nodes = Nodes + 1
    Exit Function
End If

GetLegalMoves (player) 'generates the legal moves for the current player

If LegalMovesNum > 0 Then

If depth <> 1 Then OrderMoves (player) ' move ordering to make the search faster(by producing cutoffs earlier)

'copy the legal moves array to the current procedure
For i = 1 To LegalMovesNum
LegalMovesNow(i, 1) = LegalMoves(i, 1)
LegalMovesNow(i, 2) = LegalMoves(i, 2)
Next


'store grid information in order to be able to unmake the move
For a = 1 To 8
    For b = 1 To 8
    TempGrid(a, b) = Grid(a, b)
    Next
Next
tempwhite = WhiteCount
tempblack = BlackCount

For i = 1 To LegalMovesNum
 
 Call MakeMove(player, LegalMovesNow(i, 1), LegalMovesNow(i, 2))
 
 NewScore = -Search(depth - 1, -beta, -alpha) 'calculate the score after making this move
  
'unmake the move
For a = 1 To 8
    For b = 1 To 8
    Grid(a, b) = TempGrid(a, b)
    Next
Next
 WhiteCount = tempwhite
 BlackCount = tempblack
 
 If NewScore >= beta Then
 'make a cutoff because the opponent won't let me to get in this too good position for me (he already knows a strategy to avoid this)
 Search = beta
 Exit Function
 End If
 
 
 If NewScore > alpha Then 'if the new score is better than the best one found so far
 alpha = NewScore 'make if the best score
    If depth = startdepth Then ' if we are at the root nodes then record this move because it is the best one found so far
    SelMove(1) = LegalMovesNow(i, 1)
    SelMove(2) = LegalMovesNow(i, 2)
    End If
 End If


Next

Search = alpha

Else ' if LegalMovesNum = 0

Search = -Search(depth - 1, -beta, -alpha)

End If
End Function

Private Sub OrderMoves(player As Integer)
'orders the legal moves by giving quick values for them and then sorting
Dim Values(40) As Integer
Dim TempGrid(1 To 8, 1 To 8) As Integer
Dim tempwhite As Integer, tempblack As Integer
Dim i As Integer, j As Integer, a As Integer, b As Integer
Dim temp1 As Integer, temp2 As Integer
Dim cellnum As Integer, cellval As Integer, score As Integer
Dim Counted As Integer

For i = 1 To LegalMovesNum

score = 0
Call GetFlipped(player, LegalMoves(i, 1), LegalMoves(i, 2))

'the value of a move is the sum of the values for all the discs that if flipps
For j = 1 To FlippedNum

a = NewDiscs(j, 1)
b = NewDiscs(j, 2)
 
 ' there is something very similar to this in the evaluation function
        cellnum = (b - 1) * 8 + a
        Select Case cellnum
        Case 1, 8, 57, 64
        cellval = 250
        Case 2, 7, 9, 16, 49, 56, 58, 63
        cellval = -30
        Case 10, 15, 50, 55
        cellval = -40
        Case 3, 4, 5, 6, 17, 25, 33, 41, 24, 32, 40, 48, 59, 60, 61, 62
        cellval = 15
        Case Else
        cellval = 0
        End Select
 
Counted = CountBad(a, b)
cellval = cellval - (Counted + 8 * Sgn(Counted)) ' the value of the cell also depends on how many empty cells are surrounding it
Next
 
score = score + cellval
Values(i) = score
  


Next

'sort the moves
Dim LargestVal As Integer, BestNow As Integer
For i = 1 To LegalMovesNum
LargestVal = -10000
    For j = i To LegalMovesNum
    If Values(j) > LargestVal Then BestNow = j: LargestVal = Values(j)
    Next
    temp1 = LegalMoves(BestNow, 1)
    temp2 = LegalMoves(BestNow, 2)
    LegalMoves(BestNow, 1) = LegalMoves(i, 1)
    LegalMoves(BestNow, 2) = LegalMoves(i, 2)
    LegalMoves(i, 1) = temp1
    LegalMoves(i, 2) = temp2
    Values(BestNow) = Values(i)
Next
End Sub
Private Sub ChangeTurn()
'changes the turn for the other player or passes turns or ends the game
Dim player As Integer, opponent As Integer, i As Integer, j As Integer
Dim whattodo As String
Dim PlayerMayPass As String
DrawBoard

Select Case turn
Case "black"
PlayerMayPass = "White"
player = 2
opponent = 1
Case "white"
PlayerMayPass = "Black"
player = 1
opponent = 2
End Select

whattodo = ""
For i = 1 To 8
    For j = 1 To 8
    'if there are legal moves for the othar player then only change the turn
    If IsValid(player, i, j) = True Then whattodo = "normal": Exit For
    Next
Next
'if there are no legal moves for the other player then check if we are going
'to pass the turn and give the player already moves another turn
If whattodo = "" Then

For i = 1 To 8
    For j = 1 To 8
    If IsValid(opponent, i, j) = True Then whattodo = "pass": Exit For
    Next
Next
End If

'if both players have no moves then end the game
If whattodo = "" Then whattodo = "endgame"

Select Case whattodo
Case "normal"
    Select Case turn
    Case "black"
    turn = "white"
    Case "white"
    turn = "black"
    End Select
Case "pass"
    Dim s As String
    s = PlayerMayPass + " passes"
    MsgBox (s)
    
    MovesNum = MovesNum + 1
    MovesNum1 = MovesNum
    For i = 1 To 8
    For j = 1 To 8
    GridHistory(i, j, MovesNum) = Grid(i, j)
    Next
    Next
    MoveList(MovesNum) = -1
    UpdateMoveList


Case "endgame"
    EndGame
End Select

Label2.Caption = turn
End Sub

Private Sub EndGame()
'ends the game
DrawBoard
If BlackCount > WhiteCount Then
MsgBox ("The black player wins the game")
ElseIf BlackCount < WhiteCount Then
MsgBox ("The white player wins the game")
Else
MsgBox ("The game is a draw")
End If
Timer1.Enabled = False
Timer2.Enabled = False

End Sub
Private Sub TakeMove(Index As Integer)
'takes a move for the computer or for the human player
Dim a As Integer, b As Integer, theplayer As Integer
a = (Index Mod 8) + 1
b = Int(Index / 8) + 1
Select Case turn
Case "black"
theplayer = 1
Case "white"
theplayer = 2
End Select
If IsValid(theplayer, a, b) = True Then
Call MakeMove(theplayer, a, b)
If mnusound.Checked = True Then PlaySound ("Click")
Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)

MovesNum = MovesNum + 1
MovesNum1 = MovesNum
Dim i As Integer, j As Integer
For i = 1 To 8
For j = 1 To 8
GridHistory(i, j, MovesNum) = Grid(i, j)
Next
Next
MoveList(MovesNum) = Index
UpdateMoveList

ChangeTurn
End If


DrawBoard

End Sub

Private Sub CompMove()
'makes the computer move
Dim TheDepth As Integer, best As Integer, a As Integer, b As Integer, c As Integer, t As Long

freezed = True 'make the player not able to make a move
Form1.MousePointer = 11 'change the shape of the mouse ponter
If PlayBeginner = True Then
    Dim i As Integer, player As Integer
    Dim HighestFlipped As Integer
    If turn = "black" Then player = 1 Else player = 2
    GetLegalMoves (player)
    HighestFlipped = 0
    For i = 1 To LegalMovesNum
        Call GetFlipped(player, LegalMoves(i, 1), LegalMoves(i, 2))
        Randomize
        If (FlippedNum > HighestFlipped) Or (FlippedNum = HighestFlipped And Rnd < 0.3) Then
            HighestFlipped = FlippedNum
            a = LegalMoves(i, 1): b = LegalMoves(i, 2)
        End If
        If (LegalMoves(i, 1) = 1 Or LegalMoves(i, 1) = 8) And (LegalMoves(i, 2) = 1 Or LegalMoves(i, 2) = 8) Then
            a = LegalMoves(i, 1)
            b = LegalMoves(i, 2)
            Exit For
        End If
    Next
    t = 0

Else

    Nodes = 0
    If (BlackCount + WhiteCount) > (64 - EndDepth) Then 'if we are at the end of the game then make a search to play the perfect move
        SearchEnd = True
        TheDepth = 20 'the depth is higher than the number of empty discs to handle passes
    Else 'nor at the end of the game, make a normal search
        SearchEnd = False
        TheDepth = MidDepth
    End If
    startdepth = TheDepth
    t = GetTickCount
    best = Search(TheDepth, -5000, 5000) 'make the search
    a = SelMove(1): b = SelMove(2)
    t = GetTickCount - t

End If
If (MinDelay - t) > 0 Then Sleep (MinDelay - t)
Form1.MousePointer = 1

c = 8 * (b - 1) + a - 1 'calculate the index of the square from x and y coordinates
TakeMove (c)
freezed = False
End Sub


Private Sub blackoptions_Click(Index As Integer)
Dim i As Integer
For i = 0 To 6
blackoptions(i).Checked = False
Next
blackoptions(Index).Checked = True
Player1Type = Index
Label9.Caption = blackoptions(Index).Caption
End Sub


Private Sub Command1_Click()
'the (End Setup) command button
Dim theplayer As Integer, opponent As Integer
Dim DoWhat As String
BlackCount = 0
WhiteCount = 0
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    If Grid(i, j) = 1 Then BlackCount = BlackCount + 1
    If Grid(i, j) = 2 Then WhiteCount = WhiteCount + 1
    Next
Next

Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)

If Option1.Value = True Then theplayer = 1: opponent = 2 Else: theplayer = 2: opponent = 1

DoWhat = ""

GetLegalMoves (theplayer)
If LegalMovesNum > 0 Then
DoWhat = "normal"
Select Case theplayer
Case 1
turn = "black"
Case 2
turn = "white"
End Select
Label2.Caption = turn
End If

If DoWhat = "" Then
    GetLegalMoves (opponent)
    If LegalMovesNum > 0 Then DoWhat = "pass"
End If

If DoWhat = "" Then EndGame
If DoWhat = "pass" Then ChangeTurn

MovesNum = 0
MovesNum1 = 0
For i = 1 To 8
For j = 1 To 8
GridHistory(i, j, 0) = Grid(i, j)
Next
Next

Timer1.Enabled = True
Timer2.Enabled = True
freezed = False
IsMouseDown = False
UpdateMoveList

mnusetup_Click
End Sub

Private Sub Command2_Click()
'the (clear) command button
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    Grid(i, j) = 0
    Next
Next
Grid(4, 4) = 2
Grid(5, 5) = 2
Grid(5, 4) = 1
Grid(4, 5) = 1
DrawBoard
End Sub


Private Sub Form_Load()
NewGame
Player1Type = 0
Player2Type = 3
MinDelay = 250
CommonDialog1.Flags = (cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click(Index As Integer)
If mnusetup.Checked = True Then 'if we are in the setup mode
    Dim a As Integer, b As Integer
    a = (Index Mod 8) + 1
    b = Int(Index / 8) + 1
    If Option3.Value = True Then Grid(a, b) = 1
    If Option4.Value = True Then Grid(a, b) = 2
    If Option5.Value = True Then Grid(a, b) = 0
    DrawBoard
Else ' we are in the normal mode
    If freezed = True Then Exit Sub
    TakeMove (Index)
End If

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMouseDown = True
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsMouseDown = True And mnusetup.Checked = True Then
    Dim a As Integer, b As Integer
    a = (Index Mod 8) + 1 + Int(X / 600)
    b = Int(Index / 8) + 1 + Int(Y / 600)
    If a >= 1 And b >= 1 And a <= 8 And b <= 8 Then
    If Option3.Value = True Then Grid(a, b) = 1
    If Option4.Value = True Then Grid(a, b) = 2
    If Option5.Value = True Then Grid(a, b) = 0
    End If
    DrawBoard
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMouseDown = False
End Sub

Private Sub mnudelay_Click()
On Error Resume Next
mnudelay.Checked = Not (mnudelay.Checked)
If mnudelay.Checked = True Then
Dim t As Long
restart:
t = InputBox("Enter the minimum time to delay  (in milliseconds)", "Enter time to delay", "250")
If t < 10 Or t > 5000 Then
MsgBox ("Enter a number between 10 and 5000")
GoTo restart
End If
MinDelay = t
Else
MinDelay = 0
End If
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuload_Click()
Dim i As Integer, j As Integer, DiscType As Integer, StrRow As String
CommonDialog1.Filter = "Reversi Position Files |*.pos|"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName For Input As #1

For j = 1 To 8
    Input #1, StrRow
    For i = 1 To 8
        DiscType = Val(Mid(StrRow, i, 1))
        Grid(i, j) = DiscType
    Next
Next
Input #1, turn
Close #1
DrawBoard

BlackCount = 0
WhiteCount = 0
For i = 1 To 8
    For j = 1 To 8
    If Grid(i, j) = 1 Then BlackCount = BlackCount + 1
    If Grid(i, j) = 2 Then WhiteCount = WhiteCount + 1
    Next
Next

Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)


MovesNum = 0
MovesNum1 = 0
For i = 1 To 8
For j = 1 To 8
GridHistory(i, j, 0) = Grid(i, j)
Next
Next

Timer1.Enabled = True
Timer2.Enabled = True
freezed = False
IsMouseDown = False
UpdateMoveList

End Sub

Private Sub mnunew_Click()
NewGame
End Sub

Private Sub mnuredo_Click()
If Player1Type <> 0 And Player2Type <> 0 Then Exit Sub
Dim i As Integer, j As Integer, Is2Players As Boolean
If Player1Type = 0 And Player2Type = 0 Then Is2Players = True Else Is2Players = False

Do
If Is2Players = True Then MovesNum = MovesNum + 1 Else MovesNum = MovesNum + 2
If MoveList(MovesNum + 1) <> -1 Then Exit Do
Loop

For i = 1 To 8
For j = 1 To 8
Grid(i, j) = GridHistory(i, j, MovesNum)
Next
Next

BlackCount = 0
WhiteCount = 0
For i = 1 To 8
    For j = 1 To 8
    If Grid(i, j) = 1 Then BlackCount = BlackCount + 1
    If Grid(i, j) = 2 Then WhiteCount = WhiteCount + 1
    Next
Next
Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)

DrawBoard
UpdateMoveList
End Sub

Private Sub mnusave_Click()
Dim i As Integer, j As Integer, StrRow As String
CommonDialog1.Filter = "Reversi Position Files |*.pos|"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName For Output As #1

For j = 1 To 8
    StrRow = ""
    For i = 1 To 8
        StrRow = StrRow + Trim$(Str(Grid(i, j)))
    Next
    StrRow = Trim$(StrRow)
    Write #1, StrRow
Next
Write #1, turn
Close #1
End Sub

Private Sub mnuseggest_Click()
'similar to the (CompMove) sub
If (Player1Type <> 0 And turn = "black") Or (Player2Type <> 0 And turn = "white") Then Exit Sub 'exit sub if it is computer to play
If Timer1.Enabled = False And Timer2.Enabled = False Then Exit Sub 'exit sub if the game is over
Dim OppType As Integer, imgnum As Integer
mnuseggest.Enabled = False
freezed = True
If turn = "black" Then OppType = Player2Type: imgnum = 2
If turn = "white" Then OppType = Player1Type: imgnum = 1
If OppType = 0 Then OppType = 3
Select Case OppType
Case 1
MidDepth = 1
Case 2
MidDepth = 2
Case 3
MidDepth = 4
Case 4
MidDepth = 6
End Select
EndDepth = MidDepth * 2

Dim TheDepth As Integer, best As Integer, a As Integer, b As Integer, c As Integer
Form1.MousePointer = 11
If (BlackCount + WhiteCount) > (64 - EndDepth) Then
SearchEnd = True
TheDepth = 20
Else
SearchEnd = False
TheDepth = MidDepth
End If
startdepth = TheDepth
best = Search(TheDepth, -5000, 5000)
Form1.MousePointer = 1

a = SelMove(1)
b = SelMove(2)
c = 8 * (b - 1) + a - 1

Dim i As Integer
For i = 1 To 3 'make the disc blink 3 times
    Image1(c).Picture = imgMain.ListImages.Item(imgnum).Picture
    DoEvents
    Sleep 200
    Image1(c).Picture = imgMain.ListImages.Item(4).Picture
    DoEvents
    Sleep 200
Next
freezed = False
mnuseggest.Enabled = True
End Sub

Private Sub mnusetup_Click()
mnugame.Enabled = mnusetup.Checked
mnuoptions.Enabled = mnusetup.Checked
Toolbar1.Enabled = mnusetup.Checked
Timer1.Enabled = mnusetup.Checked
Timer2.Enabled = mnusetup.Checked
mnusetup.Checked = Not (mnusetup.Checked)
Frame2.Visible = mnusetup.Checked
DrawBoard
End Sub

Private Sub mnushowpossible_Click()
mnushowpossible.Checked = Not (mnushowpossible.Checked)
If mnushowpossible.Checked = True Then Toolbar1.Buttons(9).Value = tbrPressed Else Toolbar1.Buttons(9).Value = tbrUnpressed
DrawBoard
End Sub

Private Sub mnusound_Click()
mnusound.Checked = Not (mnusound.Checked)
If mnusound.Checked = True Then Toolbar1.Buttons(10).Value = tbrPressed Else Toolbar1.Buttons(10).Value = tbrUnpressed
End Sub

Private Sub mnuundo_Click()
If Player1Type <> 0 And Player2Type <> 0 Then Exit Sub
Dim i As Integer, j As Integer, Is2Players As Boolean
If Player1Type = 0 And Player2Type = 0 Then Is2Players = True Else Is2Players = False
Dim SaveMovesnum
SaveMovesnum = MovesNum

Do
If Is2Players = True Then MovesNum = MovesNum - 1 Else MovesNum = MovesNum - 2
If MoveList(MovesNum + 1) <> -1 Or MovesNum = 0 Then Exit Do
Loop
If MovesNum < 0 Then MovesNum = SaveMovesnum: Exit Sub

For i = 1 To 8
For j = 1 To 8
Grid(i, j) = GridHistory(i, j, MovesNum)
Next
Next

If MovesNum Mod 2 = 1 Then turn = "white" Else turn = "black"
Label2.Caption = turn

BlackCount = 0
WhiteCount = 0
For i = 1 To 8
    For j = 1 To 8
    If Grid(i, j) = 1 Then BlackCount = BlackCount + 1
    If Grid(i, j) = 2 Then WhiteCount = WhiteCount + 1
    Next
Next
Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)

DrawBoard
UpdateMoveList
End Sub

Private Sub Timer1_Timer()
'make a move for the computer if the black player is a computer

'if two computers are playing then show the (Stop Aubo Play) button
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Toolbar1.Buttons(7).Enabled = True
Else
Toolbar1.Buttons(7).Enabled = False
End If

If turn = "black" And Player1Type <> 0 Then
'determine the search depth according to the computer level
If Player1Type = 2 Then PlayBeginner = True Else PlayBeginner = False
Select Case Player1Type
Case 3
MidDepth = 1
Case 4
MidDepth = 2
Case 5
MidDepth = 4
Case 6
MidDepth = 6
End Select
EndDepth = MidDepth * 2
CompMove
End If
End Sub

Private Sub Timer2_Timer()
'similar to (timer1) but for the white player
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Toolbar1.Buttons(7).Enabled = True
Else
Toolbar1.Buttons(7).Enabled = False
End If

If turn = "white" And Player2Type <> 0 Then
If Player2Type = 2 Then PlayBeginner = True Else PlayBeginner = False
Select Case Player2Type
Case 3
MidDepth = 1
Case 4
MidDepth = 2
Case 5
MidDepth = 4
Case 6
MidDepth = 6
End Select
EndDepth = MidDepth * 2
CompMove
End If
End Sub

Private Sub Timer3_Timer()
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Toolbar1.Buttons(7).Enabled = True
Else
Toolbar1.Buttons(7).Enabled = False
End If

If MovesNum < MovesNum1 Then Toolbar1.Buttons(6).Enabled = True Else Toolbar1.Buttons(6).Enabled = False
If MovesNum > 0 Then Toolbar1.Buttons(5).Enabled = True Else Toolbar1.Buttons(5).Enabled = False
'Form1.Caption = IsMouseDown
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "new"
Call mnunew_Click
Case "open"
Call mnuload_Click
Case "save"
mnusave_Click
Case "undo"
mnuundo_Click
Case "redo"
mnuredo_Click
Case "stop"
blackoptions_Click (0)
Case "show"
Call mnushowpossible_Click
Case "sound"
Call mnusound_Click
Case "help"
Call MsgBox("Sorry, help is not available yet.", vbInformation)
End Select
End Sub

Private Sub whiteoptions_Click(Index As Integer)
Dim i As Integer
For i = 0 To 6
whiteoptions(i).Checked = False
Next
whiteoptions(Index).Checked = True
Player2Type = Index
Label10.Caption = whiteoptions(Index).Caption
End Sub
