VERSION 5.00
Begin VB.Form frmSubasta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subastas"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2520
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   570
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtPrecio 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Text            =   "1"
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subastar"
      Height          =   570
      Left            =   1320
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Selecciona el item a subastar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3150
   End
   Begin VB.Label Label1 
      Caption         =   "Precio inicial"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   930
   End
End
Attribute VB_Name = "frmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

If IsNumeric(txtPrecio.Text) Then
    If guiInventoryNameGet(guiElements.inventory, "invUser", List1.listIndex + 1) <> "" Then
        Call WriteSubastar(List1.listIndex + 1, Val(txtPrecio.Text))
        Unload Me
    Else
        AddtoRichTextBox frmMain.RecTxt, "El item seleccionado es invalido.", 255, 0, 0, True, True
    End If
End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Byte
    
    For i = 1 To MAX_INVENTORY_SLOTS
        If guiInventoryNameGet(guiElements.inventory, "invUser", i) <> "" Then
            List1.AddItem guiInventoryNameGet(guiElements.inventory, "invUser", i)
        Else
            List1.AddItem "Nada"
        End If
    Next i
End Sub
