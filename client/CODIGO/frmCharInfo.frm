VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información del personaje"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharInfo.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMiembro 
      BackColor       =   &H80000006&
      Height          =   1080
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2970
      Width           =   4755
   End
   Begin VB.TextBox txtPeticiones 
      BackColor       =   &H80000006&
      Height          =   1080
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4530
      Width           =   4755
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   4215
      Top             =   5760
      Width           =   1065
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   2955
      Top             =   5775
      Width           =   1185
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   1905
      Top             =   5805
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1080
      Top             =   5790
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   90
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label ejercito 
      BackStyle       =   0  'Transparent
      Caption         =   "Faccion:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   16
      Top             =   1695
      Width           =   2880
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos asesinados:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   15
      Top             =   420
      Width           =   2850
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales asesinados:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   14
      Top             =   1185
      Width           =   2895
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3015
      TabIndex        =   13
      Top             =   900
      Width           =   2265
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   12
      Top             =   1440
      Width           =   2760
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   11
      Top             =   930
      Width           =   2880
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   10
      Top             =   1950
      Width           =   2805
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   9
      Top             =   2205
      Width           =   1950
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3015
      TabIndex        =   8
      Top             =   645
      Width           =   2130
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
      Caption         =   "Clan Actual:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3015
      TabIndex        =   7
      Top             =   1155
      Width           =   2130
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3015
      TabIndex        =   6
      Top             =   405
      Width           =   2190
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   5
      Top             =   675
      Width           =   2880
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   165
      Width           =   5640
   End
   Begin VB.Label lblSolicitado 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimas membresías solicitadas:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   315
      TabIndex        =   3
      Top             =   4215
      Width           =   2280
   End
   Begin VB.Label lblMiembro 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimos clanes en los que participó:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   315
      TabIndex        =   2
      Top             =   2655
      Width           =   2985
   End
End
Attribute VB_Name = "frmCharInfo"
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

Public frmmiembros As Boolean
Public frmsolicitudes As Boolean

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Image2_Click()
    Call WriteGuildKickMember(Right$(Nombre, Len(Nombre) - 8))
    frmmiembros = False
    frmsolicitudes = False
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub Image3_Click()
    Call WriteGuildRequestJoinerInfo(Right$(Nombre, Len(Nombre) - 8))
End Sub

Private Sub Image4_Click()
    Load frmCommet
    frmCommet.T = RECHAZOPJ
    frmCommet.Nombre = Right$(Nombre, Len(Nombre) - 8)
    frmCommet.Caption = "Ingrese motivo para rechazo"
    frmCommet.Show vbModeless, frmCharInfo
End Sub

Private Sub Image5_Click()
    frmmiembros = False
    frmsolicitudes = False
    Call WriteGuildAcceptNewMember(Trim$(Right$(Nombre, Len(Nombre) - 8)))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

