VERSION 5.00
Begin VB.Form frmRetos 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Duelo!"
      Height          =   630
      Left            =   990
      TabIndex        =   1
      Top             =   1365
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   795
      TabIndex        =   0
      Text            =   "Nombre del usuario."
      Top             =   570
      Width           =   1890
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

If Text1.Text <> " " And Text1.Text <> "" And Text1.Text <> "Nombre del Usuario." Then
    Call WriteReto
Else
    MsgBox "El nombre que ha ingresado es invalido."
End If

End Sub
