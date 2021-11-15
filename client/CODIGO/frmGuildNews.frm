VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "GuildNews"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildNews.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox news 
      Height          =   1935
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   330
      Width           =   4440
   End
   Begin VB.ListBox guerra 
      Height          =   1035
      ItemData        =   "frmGuildNews.frx":17DE5
      Left            =   390
      List            =   "frmGuildNews.frx":17DE7
      TabIndex        =   1
      Top             =   4260
      Width           =   4350
   End
   Begin VB.ListBox aliados 
      Height          =   1035
      ItemData        =   "frmGuildNews.frx":17DE9
      Left            =   390
      List            =   "frmGuildNews.frx":17DEB
      TabIndex        =   0
      Top             =   2820
      Width           =   4350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enemigos:"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   390
      TabIndex        =   5
      Top             =   4020
      Width           =   1470
   End
   Begin VB.Label lblAliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliados:"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   390
      TabIndex        =   4
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label lblNoticias 
      BackStyle       =   0  'Transparent
      Caption         =   "Noticias:"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   315
      TabIndex        =   3
      Top             =   120
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   1845
      Top             =   5700
      Width           =   1515
   End
End
Attribute VB_Name = "frmGuildNews"
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


Private Sub Image1_Click()
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub

