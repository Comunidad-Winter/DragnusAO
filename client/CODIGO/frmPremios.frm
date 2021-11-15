VERSION 5.00
Begin VB.Form frmPremios 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Premios"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   600
      Left            =   2640
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   2
      Top             =   105
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Canjear"
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   2550
      Width           =   2400
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   2385
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   90
      TabIndex        =   3
      Top             =   2160
      Width           =   2430
   End
End
Attribute VB_Name = "frmPremios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call WriteRPremios(List1.ListIndex + 1)
    
    If UserPuntos >= PremiosInv(List1.ListIndex + 1).Puntos Then
        UserPuntos = UserPuntos - PremiosInv(List1.ListIndex + 1).Puntos
        If UserPuntos >= PremiosInv(List1.ListIndex + 1).Puntos Then
            Label3.ForeColor = vbGreen
        Else
            Label3.ForeColor = vbRed
        End If
        Label3.Caption = "Puntos: " & UserPuntos & "/" & PremiosInv(List1.ListIndex + 1).Puntos
    End If
End Sub



Private Sub list1_Click()

Dim DR As RECT

DR.left = 0
DR.top = 0
DR.Right = 32
DR.bottom = 32

If UserPuntos >= PremiosInv(List1.ListIndex + 1).Puntos Then
    Label3.ForeColor = vbGreen
Else
    Label3.ForeColor = vbRed
End If

Label3.Caption = "Puntos: " & UserPuntos & "/" & PremiosInv(List1.ListIndex + 1).Puntos

'Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, PremiosInv(List1.ListIndex + 1).GrhIndex, DR, DR)

End Sub
