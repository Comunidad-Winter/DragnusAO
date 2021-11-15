VERSION 5.00
Begin VB.Form frmBancoObj 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox invUsu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   3825
      ScaleHeight     =   4005
      ScaleWidth      =   2490
      TabIndex        =   8
      Top             =   1755
      Width           =   2520
   End
   Begin VB.PictureBox InvNpc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   600
      ScaleHeight     =   4005
      ScaleWidth      =   2490
      TabIndex        =   7
      Top             =   1755
      Width           =   2520
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   405
      Left            =   3180
      TabIndex        =   5
      Text            =   "1"
      Top             =   6120
      Width           =   570
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      ScaleHeight     =   690
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   6450
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2670
      TabIndex        =   6
      Top             =   1140
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   3780
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   540
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6090
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MaxGolpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3870
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MinGolpe"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3870
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Top             =   750
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   1
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   6510
      TabIndex        =   9
      Top             =   6870
      Width           =   315
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
