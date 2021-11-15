VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6900
   ClipControls    =   0   'False
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
   Picture         =   "frmGuildDetails.frx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      Height          =   1140
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   285
      Width           =   6300
   End
   Begin VB.TextBox txtCodex1 
      Height          =   345
      Index           =   0
      Left            =   645
      TabIndex        =   7
      Top             =   2355
      Width           =   5790
   End
   Begin VB.TextBox txtCodex1 
      Height          =   330
      Index           =   1
      Left            =   630
      TabIndex        =   6
      Top             =   2835
      Width           =   5805
   End
   Begin VB.TextBox txtCodex1 
      Height          =   330
      Index           =   2
      Left            =   645
      TabIndex        =   5
      Top             =   3315
      Width           =   5790
   End
   Begin VB.TextBox txtCodex1 
      Height          =   360
      Index           =   3
      Left            =   630
      TabIndex        =   4
      Top             =   3780
      Width           =   5820
   End
   Begin VB.TextBox txtCodex1 
      Height          =   330
      Index           =   4
      Left            =   615
      TabIndex        =   3
      Top             =   4275
      Width           =   5850
   End
   Begin VB.TextBox txtCodex1 
      Height          =   345
      Index           =   5
      Left            =   630
      TabIndex        =   2
      Top             =   4755
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Height          =   345
      Index           =   6
      Left            =   630
      TabIndex        =   1
      Top             =   5235
      Width           =   5805
   End
   Begin VB.TextBox txtCodex1 
      Height          =   345
      Index           =   7
      Left            =   630
      TabIndex        =   0
      Top             =   5715
      Width           =   5820
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   1
      Left            =   4260
      Top             =   6225
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   0
      Left            =   1470
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildDetails.frx":264A8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   420
      TabIndex        =   8
      Top             =   1470
      Width           =   6255
   End
End
Attribute VB_Name = "frmGuildDetails"
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




Private Sub Form_Deactivate()

'If Not frmGuildLeader.Visible Then
'    Me.SetFocus
'Else
'    'Unload Me
'End If
'
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Dim fdesc As String
            Dim Codex() As String
            Dim k As Byte
            Dim Cont As Byte
    
            fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
            '    If Not AsciiValidos(fdesc) Then
            '        MsgBox "La descripcion contiene caracteres invalidos"
            '        Exit Sub
            '    End If

            Cont = 0
            For k = 0 To txtCodex1.UBound
            '    If Not AsciiValidos(txtCodex1(k)) Then
            '        MsgBox "El codex tiene invalidos"
            '        Exit Sub
            '    End If
                If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
            Next k
            If Cont < 4 Then
                MsgBox "Debes definir al menos cuatro mandamientos."
                Exit Sub
            End If
                        
            ReDim Codex(txtCodex1.UBound) As String
            For k = 0 To txtCodex1.UBound
                Codex(k) = txtCodex1(k)
            Next k
    
            If CreandoClan Then
                Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
            Else
                Call WriteClanCodexUpdate(fdesc, Codex)
            End If

            CreandoClan = False
            Unload Me
            
    End Select
End Sub

