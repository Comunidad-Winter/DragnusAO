VERSION 5.00
Begin VB.Form frmTorneo 
   Caption         =   "Torneo"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   2475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Descalificar"
      Height          =   465
      Left            =   420
      TabIndex        =   2
      Top             =   3000
      Width           =   1605
   End
   Begin VB.ListBox lstTorneo 
      Height          =   2010
      Left            =   210
      TabIndex        =   1
      Top             =   240
      Width           =   2040
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "Traer"
      Height          =   465
      Left            =   420
      TabIndex        =   0
      Top             =   2400
      Width           =   1605
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuitar_Click()
    If lstTorneo.listIndex > -1 Then
        Dim aux As String
        aux = mid$(general_field_read(1, lstTorneo.List(lstTorneo.listIndex), Asc("-")), 10, Len(general_field_read(1, lstTorneo.List(lstTorneo.listIndex), Asc("-"))))
        Call WriteDescalificar(aux)
        lstTorneo.RemoveItem lstTorneo.listIndex
    End If
End Sub

Private Sub cmdSum_Click()
    Dim aux As String
    aux = mid$(general_field_read(1, lstTorneo.List(lstTorneo.listIndex), Asc("-")), 10, Len(general_field_read(1, lstTorneo.List(lstTorneo.listIndex), Asc("-"))))
    Call WriteSummonChar(aux)
End Sub
