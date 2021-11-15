VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   1785
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer tTileEvents 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2955
      Top             =   540
   End
   Begin VB.Timer tCastle 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2445
      Top             =   540
   End
   Begin VB.Timer TSubasta 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3405
      Top             =   60
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4575
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer TNoche 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2415
      Top             =   60
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   480
      Top             =   60
   End
   Begin VB.Timer securityTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   960
      Top             =   60
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   540
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1020
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1920
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the message right away
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call Cerrar_Usuario(iUserIndex)
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

centinelSecs = centinelSecs + 1

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb
Call ActualizaStatsES

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long

Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        
        Horas = 0
        
    End If
    
    ContadorFacciones
    
    MinsRunning = 0
End If

    
Minutos = Minutos + 1

'�?�?�?�?�?�?�?�?�?�?�
Call ModAreas.AreasOptimizacion
'�?�?�?�?�?�?�?�?�?�?�

'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Call LimpiarMundo
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)
    Resume Next
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.toall, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call Statistics.DumpStatistics

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call closeConnection(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    Dim iNpcIndex As Integer
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser
       'Conexion activa?
       If UserList(iUserIndex).ConnID <> -1 Then
          '�User valido?
    
          If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
             
             '[Alejo-18-5]
             bEnviarStats = False
             bEnviarAyS = False
             
             UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0
    
             
             Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)
             
                    
             If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
             If UserList(iUserIndex).flags.Ceguera = 1 Or _
                UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
             
              
             If UserList(iUserIndex).flags.Muerto = 0 Then
                   
                   '[Consejeros]
                   If (UserList(iUserIndex).flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                   If UserList(iUserIndex).flags.Desnudo And (UserList(iUserIndex).flags.Privilegios And PlayerType.User) Then Call EfectoFrio(iUserIndex)
                   If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
                   If UserList(iUserIndex).flags.Envenenado = 1 And (UserList(iUserIndex).flags.Privilegios And PlayerType.User) Then Call EfectoVeneno(iUserIndex, bEnviarStats)
                   If UserList(iUserIndex).flags.AdminInvisible <> 1 Then
                        If UserList(iUserIndex).flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                        If UserList(iUserIndex).flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                   End If
                        
                   If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                    
                   Call DuracionPociones(iUserIndex)
                    
                   Call HambreYSed(iUserIndex, bEnviarAyS)
                    
                   If Lloviendo Then
                        If Not Intemperie(iUserIndex) Then
                            If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                            'No esta descansando
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                If UserList(iUserIndex).Invent.ArmourEqpObjIndex > 0 Then
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                End If
                            ElseIf UserList(iUserIndex).flags.Descansar Then
                            'esta descansando
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                'termina de descansar automaticamente
                                If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                    UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        UserList(iUserIndex).flags.Descansar = False
                                End If
                                
                            End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                        End If
                   Else
                        If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                        'No esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                            If bEnviarStats Then
                                Call WriteUpdateHP(iUserIndex)
                                bEnviarStats = False
                            End If
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                            If bEnviarStats Then
                                Call WriteUpdateSta(iUserIndex)
                                bEnviarStats = False
                            End If
                            
                        ElseIf UserList(iUserIndex).flags.Descansar Then
                        'esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                            If bEnviarStats Then
                                Call WriteUpdateHP(iUserIndex)
                                bEnviarStats = False
                            End If
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                            If bEnviarStats Then
                                Call WriteUpdateSta(iUserIndex)
                                bEnviarStats = False
                            End If
                            'termina de descansar automaticamente
                            If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                    Call WriteRestOK(iUserIndex)
                                    Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                    UserList(iUserIndex).flags.Descansar = False
                            End If
                            
                        End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                   End If
                   
                   If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                    
                   If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
           End If 'Muerto
         Else 'no esta logeado?
            'Inactive players will be removed!
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
                  UserList(iUserIndex).Counters.IdleCount = 0
                  Call closeConnection(iUserIndex)
            End If
         End If 'UserLogged
        
         'If there is anything to be sent, we send it
         Call FlushBuffer(iUserIndex)
       End If
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuCerrar_Click()
    If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. �Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        Dim f
        Dim i As Integer
        
        For i = 1 To MaxUsers
            If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
                Call closeConnection(i, True)
            Else
                If UserList(i).flags.UserLogged Then
                    Call closeChar(i)
                End If
            End If
        Next i
        
        For Each f In Forms
            Unload f
        Next
        dbDisconnect
    End If
End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo errhandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub

Private Sub securityTimer_Timer()

'Call Seguro

End Sub

Private Sub tCastle_Timer()
    Dim i As Byte
    
    For i = 1 To 2
        Castillo(i).UnderAttack = 1
    Next i

End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            ''ia comun
            If Npclist(NpcIndex).flags.Paralizado = 1 Then
                Call EfectoParalisisNpc(NpcIndex)
            Else
                'Usamos AI si hay algun user en el mapa
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                   Call EfectoParalisisNpc(NpcIndex)
                End If
                mapa = Npclist(NpcIndex).Pos.Map
                If mapa > 0 Then
                     If MapInfo(mapa).NumUsers > 0 Then
                             If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                   Call NPCAI(NpcIndex)
                             End If
                     End If
                End If
            End If
        End If
    Next NpcIndex
End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub
Private Sub tLluvia_Timer()
On Error GoTo errhandler

Dim iCount As Long
If Lloviendo Then
   For iCount = 1 To LastUser
        Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
errhandler:
Call LogError("tLluvia " & Err.Number & ": " & Err.description)
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.toall, 0, PrepareMessageRainToggle())
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.toall, 0, PrepareMessageRainToggle())
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(SendTarget.toall, 0, PrepareMessageRainToggle())
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(SendTarget.toall, 0, PrepareMessageRainToggle())
            End If
    End If
End If

Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
Static Segundos As Integer
Dim NuevaA As Boolean
Dim NuevoL As Boolean
Dim GI As Integer

Segundos = Segundos + 6

Dim i As Long

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
        If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
            UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
            Call WriteConsoleMsg(i, "Est�s obstruyendo la via p�blica, mu�vete o ser�s encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
            
            If UserList(i).Counters.PiqueteC > 23 Then
                UserList(i).Counters.PiqueteC = 0
                Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
            End If
        Else
            If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
        End If
        
        If Segundos >= 18 Then
            If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
        End If
                Call FlushBuffer(i)
    End If
    
Next i

If Segundos >= 18 Then Segundos = 0

Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call closeConnection(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call closeConnection(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub TSubasta_Timer()

MinutosSubasta = MinutosSubasta + 1
Dim i As Integer

For i = 1 To LastUser
    Call WriteConsoleMsg(i, "Van " & str(MinutosSubasta) & " minutos de la subasta y la mayor oferta hasta ahora es de " & str(MayorOferta) & ".", FontTypeNames.FONTTYPE_GUILD)
Next i

If MinutosSubasta = MaxMinutosSubasta Then
    Call TerminarSubasta
    TSubasta.Enabled = False
End If

End Sub

Private Sub tTileEvents_Timer()
On Error GoTo errhandler

    Dim i As Byte
    Dim j As Byte
    Dim b As Boolean
    
    If NumInvocaciones > 0 Then
        
        For i = 1 To NumInvocaciones
            'Solo se puede invocar cada 5 minutos.
            If Invocacion(i).EnabledCounter > 0 Then
                Invocacion(i).EnabledCounter = Invocacion(i).EnabledCounter - 1
            Else
                b = True
                For j = 1 To Invocacion(i).CantidadUsers
                    b = b And MapData(Invocacion(i).UserPos(j).Map, Invocacion(i).UserPos(j).X, Invocacion(i).UserPos(j).Y).userIndex
                    If b = False Then Exit For
                Next j
                If b Then
                    If Invocacion(i).Counter > 0 Then
                        For j = 1 To Invocacion(i).CantidadUsers
                            Call WriteConsoleMsg(MapData(Invocacion(i).UserPos(j).Map, Invocacion(i).UserPos(j).X, Invocacion(i).UserPos(j).Y).userIndex, Invocacion(i).Counter, FontTypeNames.FONTTYPE_INFO)
                        Next j
                        Invocacion(i).Counter = Invocacion(i).Counter - 1
                    Else
                        For j = 1 To Invocacion(i).CantidadNpc
                            Call SpawnNpc(Invocacion(i).NpcInvocado, Invocacion(i).NpcPos, True, False)
                        Next j
                        For j = 1 To Invocacion(i).CantidadUsers
                            Call WriteConsoleMsg(MapData(Invocacion(i).UserPos(j).Map, Invocacion(i).UserPos(j).X, Invocacion(i).UserPos(j).Y).userIndex, "Las puertas del abismo han sido abiertas.", FontTypeNames.FONTTYPE_INFO)
                        Next j
                        Invocacion(i).EnabledCounter = 180
                        'Reset counter
                        Invocacion(i).Counter = Invocacion(i).TiempoInvocacion
                    End If
                Else
                    'Reset counter
                    Invocacion(i).Counter = Invocacion(i).TiempoInvocacion
                End If
            End If
        Next i
    End If
    
    If NumTronos Then
        For i = 1 To NumTronos
            If MapData(Trono(i).Pos.Map, Trono(i).Pos.Y, Trono(i).Pos.X).userIndex Then
                If UserList(MapData(Trono(i).Pos.Map, Trono(i).Pos.Y, Trono(i).Pos.X).userIndex).GuildIndex Then
                    b = True
                Else
                    b = False
                End If
            End If
            If b Then
                Trono(i).Counter = Trono(i).Counter + 1
                If Trono(i).Counter >= Trono(i).TiempoTrono Then
                    MapInfo(Trono(i).Pos.Map).PoseidoPor = UserList(MapData(Trono(i).Pos.Map, Trono(i).Pos.Y, Trono(i).Pos.X).userIndex).GuildIndex
                Else
                    Call WriteConsoleMsg(MapData(Trono(i).Pos.Map, Trono(i).Pos.Y, Trono(i).Pos.X).userIndex, "Tomaras el mapa en " & Trono(i).TiempoTrono - Trono(i).Counter & ".", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Trono(i).Counter = 0
            End If
        Next i
    End If
    
Exit Sub
errhandler:
    Call LogError("tTileEvents " & Err.Number & ": " & Err.description)
End Sub