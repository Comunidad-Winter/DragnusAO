VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mapa SaaO"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   315
      Left            =   7110
      Top             =   45
      Width           =   330
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\GRAFICOS\Mapa.jpg")
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub
