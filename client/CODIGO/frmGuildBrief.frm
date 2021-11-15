VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7515
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
   Picture         =   "frmGuildBrief.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Desc 
      Height          =   1275
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5250
      Width           =   7035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ofrecer Paz"
      Height          =   375
      Left            =   1005
      MouseIcon       =   "frmGuildBrief.frx":2D78C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6690
      Width           =   1245
   End
   Begin VB.CommandButton aliado 
      Caption         =   "Ofrecer Alianza"
      Height          =   375
      Left            =   2250
      MouseIcon       =   "frmGuildBrief.frx":2D8DE
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6690
      Width           =   1815
   End
   Begin VB.CommandButton Guerra 
      Caption         =   "Declarar Guerra"
      Height          =   375
      Left            =   4065
      MouseIcon       =   "frmGuildBrief.frx":2DA30
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6690
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   5670
      MouseIcon       =   "frmGuildBrief.frx":2DB82
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6705
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   390
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":2DCD4
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6690
      Width           =   885
   End
   Begin VB.Label CodexTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Codex:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2775
      Width           =   1110
   End
   Begin VB.Label Reputation 
      BackStyle       =   0  'Transparent
      Caption         =   "Reputación:"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3525
      TabIndex        =   24
      Top             =   240
      Width           =   3225
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   23
      Top             =   2460
      Width           =   6975
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   22
      Top             =   2205
      Width           =   6975
   End
   Begin VB.Label lblAlineacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineacion:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   21
      Top             =   1950
      Width           =   6975
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Elecciones:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   20
      Top             =   1695
      Width           =   6975
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   19
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   18
      Top             =   1185
      Width           =   6975
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   17
      Top             =   930
      Width           =   6975
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   16
      Top             =   675
      Width           =   6480
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   15
      Top             =   420
      Width           =   3270
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   14
      Top             =   180
      Width           =   3255
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   3270
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   3015
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   3525
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   3780
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   4035
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   4290
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   4545
      Width           =   6735
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   6735
   End
End
Attribute VB_Name = "frmGuildBrief"
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

Public EsLeader As Boolean

Private Sub aliado_Click()
    frmCommet.Nombre = Right(Nombre.caption, Len(Nombre.caption) - 7)
    frmCommet.T = TIPO.ALIANZA
    frmCommet.caption = "Ingrese propuesta de alianza"
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call frmGuildSol.RecieveSolicitud(Right$(Nombre, Len(Nombre) - 7))
    Call frmGuildSol.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Command3_Click()
    frmCommet.Nombre = Right(Nombre.caption, Len(Nombre.caption) - 7)
    frmCommet.T = TIPO.PAZ
    frmCommet.caption = "Ingrese propuesta de paz"
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub Guerra_Click()
    Call WriteGuildDeclareWar(Right(Nombre.caption, Len(Nombre.caption) - 7))
    Unload Me
End Sub
