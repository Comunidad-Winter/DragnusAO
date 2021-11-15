VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":00D5
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   390
      Width           =   5505
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación del mal"
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
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación Real"
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEligeAlineacion.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Dim LastColoured As Byte

'odio programar sin tiempo (c) el oso

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDescripcion(LastColoured).BorderStyle = 0
    lblDescripcion(LastColoured).BackStyle = 0
End Sub

Private Sub lblDescripcion_Click(Index As Integer)
    Call WriteGuildFundate(Index)
    Unload Me
End Sub

Private Sub lblDescripcion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LastColoured <> Index Then
        lblDescripcion(LastColoured).BorderStyle = 0
        lblDescripcion(LastColoured).BackStyle = 0
    End If
    
    lblDescripcion(Index).BorderStyle = 1
    lblDescripcion(Index).BackStyle = 1
    
    Select Case Index
        Case 0
            lblDescripcion(Index).BackColor = &H400000
        Case 1
            lblDescripcion(Index).BackColor = &H40&
    End Select
    
    LastColoured = Index
End Sub


Private Sub lblNombre_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDescripcion(LastColoured).BorderStyle = 0
    lblDescripcion(LastColoured).BackStyle = 0
End Sub

Private Sub lblSalir_Click()
    Unload Me
End Sub
