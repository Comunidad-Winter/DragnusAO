VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form FrmLauncher 
   BorderStyle     =   0  'None
   Caption         =   "Shadows Of Angmar AO 2008"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "FrmLauncher.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7485
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   225
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   1350
      MouseIcon       =   "FrmLauncher.frx":1594
      MousePointer    =   99  'Custom
      Top             =   5835
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   1200
      MouseIcon       =   "FrmLauncher.frx":225E
      MousePointer    =   99  'Custom
      Top             =   4095
      Width           =   2070
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   1485
      MouseIcon       =   "FrmLauncher.frx":2F28
      MousePointer    =   99  'Custom
      Top             =   4905
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   1350
      MouseIcon       =   "FrmLauncher.frx":3BF2
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   1800
   End
End
Attribute VB_Name = "FrmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\GRAFICOS\launcher.jpg")
    
    Call LoadCheats

    If IsCheating Then
        Call WriteCheating
        End
    End If
    
End Sub

Private Sub Image1_Click()
    Dim VersionActual As Integer
    Dim UltimaVersion As Integer
    Image1.Visible = False
    Inet1.url = "www.ao.dveloping.com.ar/version.txt"
    'UltimaVersion = Val(Inet1.OpenURL)
    VersionActual = Val(GetVar(App.Path & "\INIT\" & "versiones.ini", "VERSION", "Val"))
    UltimaVersion = Val(Inet1.OpenURL)
    
    If UltimaVersion = 0 Then
        Call MsgBox("No se pudo comprobar la version de su cliente, se seguira con la carga del juego, pero no se garantiza su correcto funcionamiento dirijase a www.ao.dveloping.com.ar si tiene problemas.")
        Image1.Visible = True
        'Exit Sub
    End If
    
    If UltimaVersion = VersionActual Then
        Me.Visible = False
        Call Main
    Else
        If MsgBox("Su version no es la actual, ¿Desea ejecutar el AutoUpdater?.", vbYesNo) = vbYes Then
            Shell App.Path & "\AutoUpdater.exe", vbNormalFocus
            Unload Me
        Else
            Call Main
        End If
    End If
End Sub

Private Sub Image2_Click()

Call ShellExecute(0, "Open", "http://ao.dveloping.com.ar", "", App.Path, 0)

End Sub

Private Sub Image3_Click()
    Call LoadClientSetup
    frmOpciones.Show
End Sub

Private Sub Image4_Click()
    End
End Sub
