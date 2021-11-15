VERSION 5.00
Begin VB.Form frmCuenta 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "frmCuenta.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   3720
      ScaleHeight     =   915
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   7080
      Width           =   3255
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   7
      Left            =   8760
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   6
      Left            =   6120
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   5
      Left            =   3480
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   4
      Left            =   720
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   3
      Left            =   8760
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   2
      Left            =   6120
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   1
      Left            =   3480
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox charPIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   0
      Left            =   720
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image imgNewChar 
      Height          =   735
      Left            =   480
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Image imgLogChar 
      Height          =   615
      Left            =   7560
      Top             =   7440
      Width           =   2895
   End
End
Attribute VB_Name = "frmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub charPIC_Click(index As Integer)
    If index + 1 <= CharCount Then
        UserSelectedChar = index + 1
    End If
End Sub

Private Sub Form_Load()
    DrawChars
End Sub

Private Sub imgLogChar_Click()
    If CheckUserData() = True Then
        Call Login(E_MODO_LOGIN.CharLogin)
    End If
End Sub

Private Sub imgNewChar_Click()
    frmCuenta.visible = False
    frmCrearPersonaje.Show
End Sub

Public Sub DrawChars()
    Dim i As Integer
    
    Dim iBodyGrh As Integer
    Dim iHeadOffsetY As Integer
    
    Dim iWeaponGrh As Integer
    Dim iShieldGrh As Integer
    Dim iHeadGrh As Integer
    Dim iHelmGrh As Integer
    
    
    For i = 1 To CharCount
        Call Engine.getBodyData(UserChars(i).iBody, iBodyGrh, iHeadOffsetY)
        Call Engine.getHeadData(UserChars(i).iHead, iHeadGrh)
        Call Engine.getWeaponData(UserChars(i).iWeapon, iWeaponGrh)
        Call Engine.getShieldData(UserChars(i).iShield, iShieldGrh)
        Call Engine.getHelmData(UserChars(i).iHelm, iHelmGrh)
        If iBodyGrh Then _
            Grh_Render_To_Hdc iBodyGrh, charPIC(i - 1).hdc, 25, 50, True, True
        If iHeadGrh Then _
            Grh_Render_To_Hdc iHeadGrh, charPIC(i - 1).hdc, 25 + 4, 48 + iHeadOffsetY, True, True
        If iHelmGrh Then _
            Grh_Render_To_Hdc iHelmGrh, charPIC(i - 1).hdc, 25 + 4, 48 + iHeadOffsetY, True, True
        If iWeaponGrh Then _
            Grh_Render_To_Hdc iWeaponGrh, charPIC(i - 1).hdc, 25, 50, True, True
        If iShieldGrh Then _
            Grh_Render_To_Hdc iShieldGrh, charPIC(i - 1).hdc, 25, 50, True, True
    Next i
End Sub

Private Sub Picture1_Click()
    End
End Sub
