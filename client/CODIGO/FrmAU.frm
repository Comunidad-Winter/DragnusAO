VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmAU 
   BackColor       =   &H80000012&
   Caption         =   "SaaO AutoUpdater"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   Picture         =   "FrmAU.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   30
      Top             =   2595
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblprogress 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   3135
      TabIndex        =   1
      Top             =   2280
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   1860
      Top             =   3300
      Width           =   2145
   End
   Begin VB.Label lblAct 
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1635
      TabIndex        =   0
      Top             =   1680
      Width           =   3390
   End
End
Attribute VB_Name = "FrmAU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private zip As ZipExtractionClass
Private ByteArray() As Byte
Private Cuantas As Integer
Private LastVersion As Integer
Private VersionActual As Integer

Private Sub Extraer()
'EXTRAE
Dim ZipName As String
Set zip = New ZipExtractionClass
ZipName = VersionActual & ".zip"
   
If zip.OpenZip(App.Path & "\INIT\" & ZipName) Then
    Call zip.Extract(App.Path, False, True)
    Call zip.CloseZip
    If FileExist(App.Path & "\GRAFICOS.zip", vbArchive) Then
       Call zip.OpenZip(App.Path & "\GRAFICOS.zip")
       Call zip.Extract(App.Path, True, True)
       Call zip.CloseZip
    End If
    If FileExist(App.Path & "\INIT.zip", vbArchive) Then
        Call zip.OpenZip(App.Path & "\INIT.zip")
        Call zip.Extract(App.Path, True, True)
        Call zip.CloseZip
    End If
    If FileExist(App.Path & "\MAPAS.zip", vbArchive) Then
        Call zip.OpenZip(App.Path & "\MAPAS.zip")
        Call zip.Extract(App.Path, True, True)
        Call zip.CloseZip
        Debug.Print "lo baja"
    End If
    If FileExist(App.Path & "\MIDI.zip", vbArchive) Then
        Call zip.OpenZip(App.Path & "\MIDI.zip")
        Call zip.Extract(App.Path, True, True)
        Call zip.CloseZip
    End If
    If FileExist(App.Path & "\WAV.zip", vbArchive) Then
        Call zip.OpenZip(App.Path & "\WAV.zip")
        Call zip.Extract(App.Path, True, True)
        Call zip.CloseZip
    End If
    If FileExist(App.Path & "\AppPath.zip", vbArchive) Then
        Call zip.OpenZip(App.Path & "\AppPath.zip")
        Call zip.Extract(App.Path, False, True)
        Call zip.CloseZip
    End If
    
End If
   
Set zip = Nothing
   
End Sub

Private Sub CheckLastVersion()

Inet1.URL = "www.ao.dveloping.com.ar/version.txt"
LastVersion = Val(Inet1.OpenURL)

If LastVersion = 0 Then
    Call MsgBox("Se ha producido un error al comprobar actualizaciones, www.shadowsofangmar.com.ar")
    Exit Sub
End If

VersionActual = Val(GetVar(App.Path & "\INIT\versiones.ini", "VERSION", "Val"))

If LastVersion > VersionActual Then
    Cuantas = LastVersion - VersionActual
    lblprogress.Caption = Cuantas
Else
    Cuantas = 0
    lblprogress.Caption = "0"
End If

lblAct.Caption = "Preparado."

End Sub

Private Sub Download()
'Descarga de la actualizacion

On Error Resume Next

Dim Fichero As String
Dim nFic As Integer
Dim strURL As String

Fichero = App.Path & "\INIT\" & (VersionActual + 1) & ".zip"
Inet1.AccessType = icDefault

lblAct.ForeColor = vbYellow
lblAct.Caption = "Descargando..."

While Cuantas > 0
    DoEvents 'Procesamos eventos.
    lblprogress.Caption = Cuantas
    strURL = "www.ao.dveloping.com.ar/" & (VersionActual + 1) & ".zip"
    ByteArray() = Inet1.OpenURL(strURL, icByteArray)
    nFic = FreeFile
    Open Fichero For Binary Access Write As #nFic
        Put #nFic, , ByteArray()
    Close #nFic
    VersionActual = VersionActual + 1
    Call Extraer
    Kill App.Path & "\GRAFICOS.zip"
    Kill App.Path & "\INIT.zip"
    Kill App.Path & "\MAPAS.zip"
    Kill App.Path & "\MIDI.zip"
    Kill App.Path & "\WAV.zip"
    Kill App.Path & "\AppPath.zip"
    Cuantas = Cuantas - 1
    lblprogress.Caption = Cuantas
    Kill App.Path & "\INIT\" & VersionActual & ".zip"
    Fichero = App.Path & "\INIT\" & (VersionActual + 1) & ".zip"
Wend


lblAct.Caption = "Finalizado."
lblprogress = Cuantas

If MsgBox("¿Desea ejecutar el juego?", vbYesNo) Then
    Shell App.Path & "\SAAO AB.exe", vbNormalFocus
    End
Else
    End
End If

End Sub

Private Sub Form_Load()
    Call CheckLastVersion
End Sub

Private Sub Image1_Click()
    
    If Cuantas > 0 Then
        Call Download
    Else
        Call MsgBox("Ya tienes la ultima version.")
    End If
    
End Sub
