VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCrearPersonaje.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":0CCA
      Left            =   7095
      List            =   "frmCrearPersonaje.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   7200
      Width           =   2985
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":0CCE
      Left            =   7095
      List            =   "frmCrearPersonaje.frx":0CD8
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   6240
      Width           =   2985
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frmCrearPersonaje.frx":0CEB
      Left            =   7095
      List            =   "frmCrearPersonaje.frx":0CED
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5280
      Width           =   2985
   End
   Begin VB.ComboBox lstHogar 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0CEF
      Left            =   6180
      List            =   "frmCrearPersonaje.frx":0CF1
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4020
      Width           =   1860
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   4395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   10050
      TabIndex        =   33
      Top             =   3330
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4170
      TabIndex        =   32
      Top             =   2610
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":0CF3
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":0E45
      MousePointer    =   99  'Custom
      Top             =   3810
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":0F97
      MousePointer    =   99  'Custom
      Top             =   4050
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":10E9
      MousePointer    =   99  'Custom
      Top             =   4260
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":123B
      MousePointer    =   99  'Custom
      Top             =   4500
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":138D
      MousePointer    =   99  'Custom
      Top             =   4710
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":14DF
      MousePointer    =   99  'Custom
      Top             =   4950
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":1631
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":1783
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":18D5
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":1A27
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":1B79
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":1CCB
      MousePointer    =   99  'Custom
      Top             =   6300
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":1E1D
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   4560
      MouseIcon       =   "frmCrearPersonaje.frx":1F6F
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   4560
      MouseIcon       =   "frmCrearPersonaje.frx":2C39
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":2D8B
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   4560
      MouseIcon       =   "frmCrearPersonaje.frx":2EDD
      MousePointer    =   99  'Custom
      Top             =   4050
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   4560
      MouseIcon       =   "frmCrearPersonaje.frx":302F
      MousePointer    =   99  'Custom
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":3181
      MousePointer    =   99  'Custom
      Top             =   4500
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":3E4B
      MousePointer    =   99  'Custom
      Top             =   4710
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   4620
      MouseIcon       =   "frmCrearPersonaje.frx":3F9D
      MousePointer    =   99  'Custom
      Top             =   4950
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":40EF
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":4241
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":4393
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":44E5
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":51AF
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":5301
      MousePointer    =   99  'Custom
      Top             =   6300
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":5453
      MousePointer    =   99  'Custom
      Top             =   6510
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":55A5
      MousePointer    =   99  'Custom
      Top             =   6540
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":56F7
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":5849
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":599B
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":5AED
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":5C3F
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":5D91
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":5EE3
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   3750
      MouseIcon       =   "frmCrearPersonaje.frx":6035
      MousePointer    =   99  'Custom
      Top             =   7410
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":6187
      MousePointer    =   99  'Custom
      Top             =   7650
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":62D9
      MousePointer    =   99  'Custom
      Top             =   7650
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   4590
      MouseIcon       =   "frmCrearPersonaje.frx":642B
      MousePointer    =   99  'Custom
      Top             =   7890
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   3780
      MouseIcon       =   "frmCrearPersonaje.frx":657D
      MousePointer    =   99  'Custom
      Top             =   7890
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   1230
      Index           =   2
      Left            =   9630
      MouseIcon       =   "frmCrearPersonaje.frx":66CF
      MousePointer    =   99  'Custom
      Top             =   1740
      Width           =   1290
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   1
      Left            =   6210
      MouseIcon       =   "frmCrearPersonaje.frx":6821
      MousePointer    =   99  'Custom
      Top             =   8190
      Width           =   1110
   End
   Begin VB.Image boton 
      Height          =   495
      Index           =   0
      Left            =   9060
      MouseIcon       =   "frmCrearPersonaje.frx":6973
      MousePointer    =   99  'Custom
      Top             =   8190
      Width           =   2130
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4080
      TabIndex        =   28
      Top             =   7860
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   4080
      TabIndex        =   27
      Top             =   7620
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   4080
      TabIndex        =   26
      Top             =   7380
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   4080
      TabIndex        =   25
      Top             =   7170
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   24
      Top             =   6960
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   23
      Top             =   6720
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   22
      Top             =   6510
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   21
      Top             =   6270
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   20
      Top             =   6030
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   19
      Top             =   5820
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   18
      Top             =   5610
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   17
      Top             =   5400
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   16
      Top             =   5160
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   15
      Top             =   4920
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   14
      Top             =   4710
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   13
      Top             =   4500
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   12
      Top             =   4260
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   4020
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   10
      Top             =   3810
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   9
      Top             =   3330
      Width           =   270
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   3570
      Width           =   270
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   8970
      TabIndex        =   6
      Top             =   2700
      Width           =   300
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9900
      TabIndex        =   5
      Top             =   3330
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   8970
      TabIndex        =   4
      Top             =   2220
      Width           =   300
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   8970
      TabIndex        =   3
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   8970
      TabIndex        =   2
      Top             =   1770
      Width           =   300
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   8970
      TabIndex        =   1
      Top             =   1350
      Width           =   300
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(index As Integer)
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT
    
    Select Case index
        Case 0
            
            Dim i As Integer
            Dim k As Object
            i = 1
            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next
            
            UserName = txtNombre.Text
            
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
            End If
            
            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex + 1
            UserClase = lstProfesion.ListIndex + 1
            
            UserAtributos(1) = Val(lbFuerza.Caption)
            UserAtributos(2) = Val(lbInteligencia.Caption)
            UserAtributos(3) = Val(lbAgilidad.Caption)
            UserAtributos(4) = Val(lbCarisma.Caption)
            UserAtributos(5) = Val(lbConstitucion.Caption)
            
            UserHogar = lstHogar.ListIndex + 1
            
            'Barrin 3/10/03
            If CheckData() Then
                frmPasswd.Show vbModal, Me
            End If
            
        Case 1
            modSound.Music_Play ("2.mid")
            Me.Visible = False
            frmCuenta.Show
        Case 2
            modSound.Sound_Play SND_DICE, DSBPLAY_DEFAULT
            Call TirarDados
    End Select
End Sub

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub Command1_Click(index As Integer)
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT
    
    Dim indice
    If index Mod 2 = 0 Then
        If SkillPoints > 0 Then
            indice = index \ 2
            Skill(indice).Caption = Val(Skill(indice).Caption) + 1
            SkillPoints = SkillPoints - 1
        End If
    Else
        If SkillPoints < 10 Then
            
            indice = index \ 2
            If Val(Skill(indice).Caption) > 0 Then
                Skill(indice).Caption = Val(Skill(indice).Caption) - 1
                SkillPoints = SkillPoints + 1
            End If
        End If
    End If
    
    Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(DirGraficos & "CP-Interface.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstHogar.Clear

For i = LBound(Ciudades()) To UBound(Ciudades())
    lstHogar.AddItem Ciudades(i)
Next i


lstRaza.Clear

For i = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(i)
Next i


lstProfesion.Clear

For i = LBound(ListaClases()) To UBound(ListaClases())
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 1

Call TirarDados
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    
'TODO : Esto vuela en la 0.12.1!!!
    If lstProfesion.ListIndex + 1 = eClass.Druid Then
        Call MsgBox("Esta clase se encuentra deshabilitada hasta el próximo parche, en el que se le realizarán varios cambios importantes." & vbCrLf _
            & "Sepan disculpar las molestias.")
        
        lstProfesion.ListIndex = 0
    End If
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
