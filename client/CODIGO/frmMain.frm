VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   8715
   ClientLeft      =   360
   ClientTop       =   270
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6240
      Top             =   4005
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tAntiEngine 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3780
      Top             =   3435
   End
   Begin VB.Timer tCheat 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3810
      Top             =   4530
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   10485
      MouseIcon       =   "frmMain.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5430
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   9180
      MouseIcon       =   "frmMain.frx":0A1C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5430
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   8970
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   8940
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4830
      Top             =   2475
   End
   Begin VB.Timer ATecho 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2070
      Top             =   2490
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1965
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   2520
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6615
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   195
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5760
      Top             =   3255
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5625
      Top             =   4335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   2580
      Top             =   3780
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3120
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1590
      Left            =   195
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   210
      Visible         =   0   'False
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   2805
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0B6E
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
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   195
      ScaleHeight     =   386
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   542
      TabIndex        =   28
      Top             =   2505
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Image ExpPic 
      Height          =   90
      Left            =   9765
      Picture         =   "frmMain.frx":0BEB
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   1920
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8790
      MouseIcon       =   "frmMain.frx":0F45
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   2580
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10155
      MouseIcon       =   "frmMain.frx":1097
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2580
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image cmdInfo 
      Height          =   390
      Left            =   10455
      MouseIcon       =   "frmMain.frx":11E9
      MousePointer    =   99  'Custom
      Top             =   5490
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image CmdLanzar 
      Height          =   390
      Left            =   9000
      MouseIcon       =   "frmMain.frx":133B
      MousePointer    =   99  'Custom
      Top             =   5475
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   10695
      TabIndex        =   24
      Top             =   8610
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   10155
      TabIndex        =   23
      Top             =   8310
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblEscu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Escu"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   4125
      TabIndex        =   22
      Top             =   8565
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblHit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   750
      TabIndex        =   21
      Top             =   8565
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblArmour 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   1845
      TabIndex        =   20
      Top             =   8565
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCasco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Casco"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2940
      TabIndex        =   19
      Top             =   8565
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8775
      TabIndex        =   18
      Top             =   7950
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblMAN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8775
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8775
      TabIndex        =   16
      Top             =   8460
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8775
      TabIndex        =   15
      Top             =   6450
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSTA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8775
      TabIndex        =   13
      Top             =   7455
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image AGUApic 
      Height          =   150
      Left            =   8700
      Picture         =   "frmMain.frx":148D
      Stretch         =   -1  'True
      Top             =   8490
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image COMIDApic 
      Height          =   150
      Left            =   8685
      Picture         =   "frmMain.frx":4861
      Stretch         =   -1  'True
      Top             =   7980
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image MANpic 
      Height          =   150
      Left            =   8685
      Picture         =   "frmMain.frx":7C37
      Stretch         =   -1  'True
      Top             =   6990
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image HPpic 
      Height          =   150
      Left            =   8685
      Picture         =   "frmMain.frx":B011
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11700
      Top             =   -15
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   11415
      Top             =   15
      Width           =   270
   End
   Begin VB.Label label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   9150
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   195
      Index           =   2
      Left            =   10860
      Top             =   8640
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   195
      Index           =   1
      Left            =   11325
      Top             =   8655
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   270
      Index           =   0
      Left            =   10245
      Top             =   6705
      Width           =   1350
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   10455
      TabIndex        =   12
      Top             =   1470
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   180
      Index           =   1
      Left            =   11250
      MouseIcon       =   "frmMain.frx":E3F7
      MousePointer    =   99  'Custom
      Top             =   3210
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   195
      Index           =   0
      Left            =   11250
      MouseIcon       =   "frmMain.frx":E549
      MousePointer    =   99  'Custom
      Top             =   3570
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Left            =   9795
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   8955
      TabIndex        =   10
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10455
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9999999999"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10425
      TabIndex        =   4
      Top             =   6720
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   10185
      MouseIcon       =   "frmMain.frx":E69B
      MousePointer    =   99  'Custom
      Top             =   8250
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   1
      Left            =   10200
      MouseIcon       =   "frmMain.frx":E7ED
      MousePointer    =   99  'Custom
      Top             =   7665
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   0
      Left            =   10200
      MouseIcon       =   "frmMain.frx":E93F
      MousePointer    =   99  'Custom
      Top             =   7125
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image PicMH 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10785
      Picture         =   "frmMain.frx":EA91
      Stretch         =   -1  'True
      Top             =   8595
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,00,00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   6375
      TabIndex        =   2
      Top             =   8565
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image PicSeg 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10350
      Picture         =   "frmMain.frx":F8A3
      Stretch         =   -1  'True
      Top             =   8580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image STApic 
      Height          =   150
      Left            =   8685
      Picture         =   "frmMain.frx":FD5B
      Stretch         =   -1  'True
      Top             =   7485
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ItemName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9165
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image InvEqu 
      Height          =   3705
      Left            =   8610
      Top             =   2400
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public tX As Integer
Public tY As Integer
Public mouseX As Long
Public mouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public FirstTime As Boolean

'Dim gDSB As DirectSoundBuffer
'Dim gD As DSBUFFERDESC
'Dim gW As WAVEFORMATEX
'Dim gFileName As String
'Dim dsE As DirectSoundEnum
'Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim PuedeMacrear As Boolean




Private Sub ATecho_Timer()
#If ConAlfaB Then
'If bTecho Then
    
'    If (AlphaTechos - 5) > 0 Then
'        AlphaTechos = AlphaTechos - 5
'    Else
'        AlphaTechos = 0
'    End If
    
'Else

'    If (AlphaTechos + 10) < 255 Then
'        AlphaTechos = AlphaTechos + 10
'    Else
'        AlphaTechos = 255
'    End If
    
'End If

ATecho.Enabled = False
#End If
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
    If hlst.listIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case index
        Case 1 'subir
            If hlst.listIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.listIndex = hlst.listCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(index, hlst.listIndex + 1)
    
    Select Case index
        Case 1 'subir
            sTemp = hlst.List(hlst.listIndex - 1)
            hlst.List(hlst.listIndex - 1) = hlst.List(hlst.listIndex)
            hlst.List(hlst.listIndex) = sTemp
            hlst.listIndex = hlst.listIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.listIndex + 1)
            hlst.List(hlst.listIndex + 1) = hlst.List(hlst.listIndex)
            hlst.List(hlst.listIndex) = sTemp
            hlst.listIndex = hlst.listIndex + 1
    End Select
End Sub

Public Sub ActivarMacroHechizos()
    If Not hlst.visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, False)
        Exit Sub
    End If
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, False)
    PicMH.visible = True
End Sub

Public Sub DesactivarMacroHechizos()
        PicMH.visible = False
        TrainingMacro.Enabled = False
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, False)
End Sub
Public Sub DibujarMH()
PicMH.visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    guiInputKey KeyAscii
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
    Game_KeyEvents keyCode, Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Image2_Click()
    frmMain.WindowState = vbMinimized
End Sub

Private Sub Image4_Click()
    prgRun = False
End Sub



Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub macrotrabajo_Timer()
    If guiInventoryItemGet(guiElements.inventory, "invUser") Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not modApi.IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If (UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PanelDer_Click()

End Sub

Private Sub PicMH_Click()
    AddtoRichTextBox frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
End Sub

Private Sub PicSeg_Click()
    AddtoRichTextBox frmMain.RecTxt, "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub

Private Sub RecTxt_Click()
    If picInv.visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub RecTxt_GotFocus()
    If picInv.visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_KeyUp(keyCode As Integer, Shift As Integer)
    'Send text
    If keyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        keyCode = 0
        SendTxt.visible = False
    End If
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''







Private Sub tAntiEngine_Timer()
Dim Tick As Long
Static Times As Byte

'Agregue el contador de veces para que no te saque a cada rato si te anduvo mas lento.

If Not FirstTime Then
    FirstTime = True
    TiempoActual = GetTickCount ' we start counting here.
    Exit Sub
End If

Tick = GetTickCount

If Tick - TiempoActual > 5000 Then
    If Times > 2 Then
        MsgBox ("Se ha cerrado el juego debido al posible uso de cheats, reloguee.")
        Call WriteCheating
        End
    Else
        Times = Times + 1
    End If
Else
    Times = 0
End If

TiempoActual = GetTickCount

End Sub

Private Sub tCheat_Timer()
    Static Mins As Integer

    If Mins >= 3 Then
        Mins = 0
        If IsCheating Then
            Call WriteCheating
            'Cerramos?
            End
        End If
    Else
        Mins = Mins + 1
    End If
    
End Sub

Private Sub Timer1_Timer()
    Call WriteAceptaONo(0)
    frmMain.Timer1.Enabled = False
End Sub


''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not modApi.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If hlst.List(hlst.listIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.listIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
    
    'Call ConvertCPtoTP(MainViewShp.left, MainViewShp.top, MouseX, MouseY, tX, tY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.listIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        Call WriteCastSpell(hlst.listIndex + 1)
        Call WriteWork(eSkill.Magia)
        UsaMacro = True
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    Call WriteSpellInfo(hlst.listIndex + 1)
End Sub

Private Sub DespInv_Click(index As Integer)
    'Inventario.ScrollInventory (index = 0)
End Sub

Private Sub Form_Load()
    
    'frmMain.Caption = "Argentum Online" & " V " & App.Major & "." & _
    App.Minor & "." & App.Revision
    
    'Borre el panel derecho xP
    'PanelDer.Picture = LoadPicture(App.Path & _
    '"\Graficos\Principalnuevo_sin_energia.jpg")
    
    'RecTxt.Refresh
    
   Me.left = 0
   Me.top = 0
   
End Sub

Private Sub hlst_KeyDown(keyCode As Integer, Shift As Integer)
       keyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(keyCode As Integer, Shift As Integer)
        keyCode = 0
End Sub

Private Sub Image1_Click(index As Integer)
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT

    Select Case index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            initStatLabels
            GUIShowForm guiElements.frmEstadisticas
            LlegaronAtrib = False
            LlegaronSkills = False
        
        Case 2
            Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub Image3_Click(index As Integer)
    Select Case index
        Case 0
            'Inventario.SelectGold
            If UserGLD > 0 Then
                'frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim i As Integer
    guiListClear guiElements.frmSkills3, "lstSkills"
    For i = 1 To NUMSKILLS
        guiListAddItem guiElements.frmSkills3, "lstSkills", userSkills(i)
    Next i
    Alocados = SkillPoints
    
    guiLabelCaptionSet guiElements.frmSkills3, "lblLibres", "Puntos: " & SkillPoints

    GUIShowForm guiElements.frmSkills3
End Sub

Private Sub Label4_Click()
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT

    InvEqu.picture = LoadPicture(DirGraficos & "Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.visible = True

    hlst.visible = False
    cmdINFO.visible = False
    CmdLanzar.visible = False
    
    cmdMoverHechi(0).visible = True
    cmdMoverHechi(1).visible = True
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    Render_Inventory = True
    
    'lblItemName.Visible = True
End Sub

Private Sub Label7_Click()
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT

    InvEqu.picture = LoadPicture(DirGraficos & "Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.visible = False
    hlst.visible = True
    cmdINFO.visible = True
    CmdLanzar.visible = True
    
    cmdMoverHechi(0).visible = True
    cmdMoverHechi(1).visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    'lblItemName.Visible = False
End Sub

Private Sub picInv_DblClick()
    'If frmCarp.visible Or frmHerrero.visible Or hlst.visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then _
                     DesactivarMacroTrabajo
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT
End Sub


Private Sub RecTxt_KeyDown(keyCode As Integer, Shift As Integer)
    If picInv.visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(Mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(keyCode As Integer, Shift As Integer)
    'Send text
    If keyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        keyCode = 0
        Me.SendCMSTXT.visible = False
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(Mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.visible = False
    frmCuenta.visible = False
    frmCrearPersonaje.visible = False
    frmPasswd.visible = False
    frmConnect.visible = True
    
    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        userSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    Call Engine.Char_Remove_All
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
    
    'Call SetMusicInfo("", "", "", "Games", "{1}{0}")
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    'If frmOldPersonaje.Visible Then
    '    frmOldPersonaje.Visible = False
    'End If
    
    If Not frmCrearPersonaje.visible Then
       frmConnect.Show
    Else
        'frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Public Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If Engine.Map_Char_Get(tX, tY) > 0 Then
        If Engine.Char_Invisible_Get(Engine.Map_Char_Get(tX, tY)) = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If Engine.Char_Label_Get(Engine.Map_Char_Get(tX, tY)) <> "" Then
                m.ListaSetItem 0, Engine.Char_Label_Get(Engine.Map_Char_Get(tX, tY)), True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If Not frmPasswd.visible And Not frmCrearPersonaje.visible Then
        frmConnect.visible = True
        frmConnect.MouseIcon = LoadPicture(App.Path & "\GRAFICOS\Icons\Espada.ico")
    End If
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name And Forms(i).name <> frmPasswd.name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        userSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((Mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (Mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = Mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True

    Call Login
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    'If frmOldPersonaje.Visible Then
    '    frmOldPersonaje.Visible = False
    'End If

    If Not frmCrearPersonaje.visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If

Public Sub socketConnect()
    'CONECTAR
    #If UsarWrench = 1 Then
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
    #Else
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If
End Sub

Public Sub socketDisconnect()
    'CERRAR POSIBLES SOCKETS ABIERTOS
    #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
    #Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
    #End If
End Sub
