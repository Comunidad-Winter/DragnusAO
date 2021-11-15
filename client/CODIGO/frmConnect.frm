VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmConnect 
   BorderStyle     =   0  'None
   Caption         =   "Shadows Of Angmar AO 2008"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   9030
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1290
      TabIndex        =   4
      Text            =   "7666"
      Top             =   945
      Visible         =   0   'False
      Width           =   1875
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   225
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4005
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4125
      Width           =   3900
   End
   Begin VB.TextBox NameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4005
      TabIndex        =   2
      Top             =   3420
      Width           =   3900
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3195
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   2
      Left            =   11610
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   840
      Index           =   1
      Left            =   3975
      Top             =   6240
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   1125
      Index           =   0
      Left            =   3990
      Top             =   4785
      Width           =   3960
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

'Private Sub downloadServer_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el código de sus servidor de esta forma.
'Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los
'cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo,
'no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.
'***********************************
    'Call ShellExecute(0, "Open", "http://sourceforge.net/project/downloading.php?group_id=67718&use_mirror=osdn&filename=AOServerSrc.zip&86289150", "", App.Path, 0)
'End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
If keyCode = 27 Then
    Call SaveGameini
    prgRun = False
End If
End Sub

Private Sub Form_Load()
    Me.picture = LoadPicture(App.Path & "\Resources\Graphics\" & "Conectar.jpg")
    PortTxt.Text = Config_Inicio.Puerto
    version.caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
End Sub
Private Sub Image1_Click(index As Integer)
    modSound.Sound_Play SND_CLICK, DSBPLAY_DEFAULT

    Select Case index
        Case 0
            'Update user info.
            UserName = NameTxt.Text
            UserPassword = PasswordTxt.Text
            
            frmMain.socketDisconnect
            
            frmMain.socketConnect
            
            Call Login(E_MODO_LOGIN.AccountLogin)
        Case 2
            Call SaveGameini
            frmConnect.MousePointer = 1
            frmMain.MousePointer = 1
            prgRun = False
    End Select
End Sub
