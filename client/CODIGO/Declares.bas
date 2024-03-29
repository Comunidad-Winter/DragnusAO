Attribute VB_Name = "modDeclaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

'Objetos p�blicos
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New clsDialogos

'Tile Engine
Public Engine As New clsTileEngineX

'Meteorologic
Public Meteo As New clsMeteo

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

''
'The main timer of the game.
Public MainTimer As New clsTimer

Public TiempoActual As Long

'Sonidos
Public Const SND_CLICK As Byte = 210
Public Const SND_PASOS1 As Byte = 23
Public Const SND_PASOS2 As Byte = 24
Public Const SND_NAVEGANDO As Byte = 50
Public Const SND_OVER As Byte = 211
Public Const SND_DICE As Byte = 212

' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2788
Public Const INT_MACRO_TRABAJO As Integer = 900

Public Const INT_ATTACK As Integer = 1400
Public Const INT_ARROWS As Integer = 1000
Public Const INT_CAST_SPELL As Integer = 1000
Public Const INT_WORK As Integer = 700
Public Const INT_USEITEMU As Integer = 600
Public Const INT_USEITEMDCK As Integer = 220
Public Const INT_SENTRPU As Integer = 2000
Public Const INT_ATTACKSPELL As Integer = 1400

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500

Public Const NUMATRIBUTES As Byte = 5

Public Const URL_NOTICIAS As String = "www.ao.dveloping.com.ar/txtNoticias.txt"

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String

Public Type tColor
    R As Byte
    G As Byte
    B As Byte
End Type

Public ColoresPJ(0 To 50) As tColor


Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public Type arrayHerrero
    index As Integer
    name As String
End Type

Public Enum craftingObj
    weapon
    shield
    helmet
    armor
End Enum

Public craftingBlackSmith As Byte
Public ArmasHerrero(0 To 100) As arrayHerrero
Public ArmadurasHerrero(0 To 100) As arrayHerrero
Public CascosHerrero(0 To 100) As arrayHerrero
Public EscudosHerrero(0 To 100) As arrayHerrero
Public ObjCarpintero(0 To 100) As Integer

'User Guild Info
Public guilds() As String
Public members() As String
Public solicitudes() As String
Public paz() As String

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte

'TCP FLAGS
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As inventory
'[/KEVIN]

Public PremiosInv(1 To 20) As PremiosList

Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 20
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const MAXSKILLPOINTS As Byte = 100

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const FOgata As Integer = 1521


Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Cl�rigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladr�n
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Palad�n
    Hunter      'Cazador
    Fisher      'Pescador
    Blacksmith  'Herrero
    Lumberjack  'Le�ador
    Miner       'Minero
    Carpenter   'Carpintero
    Pirat       'Pirata
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
    Orco
End Enum

Public Enum eSkill
    Suerte = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    armas = 5
    Meditar = 6
    Apu�alar = 7
    Ocultarse = 8
    Supervivencia = 9
    Talar = 10
    Comerciar = 11
    Defensa = 12
    Pesca = 13
    Mineria = 14
    Carpinteria = 15
    Herreria = 16
    Liderazgo = 17
    Domar = 18
    Proyectiles = 19
    Wresterling = 20
    Navegacion = 21
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLe�a = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otCualquiera = 1000
End Enum

'Path const
Public Const PATH_GRAPHICS  As String = "\graphics"
Public Const PATH_MAPS  As String = "\maps"
Public Const PATH_SOUNDS  As String = "\sounds"
Public Const PATH_SCRIPTS  As String = "\scripts"
Public Const PATH_INIT  As String = "\init"
Public Const PATH_RESOURCES As String = "\resources"
Public Const SfxPath As String = "\wav"
Public Const MusicPath As String = "\midi"
'Path vars
Public resource_path As String

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "��Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertir�s en uno de ellos y ser�s perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "�Est�s meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_GOLPE_CABEZA As String = "��La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "��La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "��La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "��La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "��La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "��La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "��"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "��Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "��Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el �rbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Public Type inventory
    objIndex As Integer
    name As String
    grhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    amount As Long
    '[/Alejo]
    equipped As Byte
    valor As Long
    objType As Integer
    def As Integer
    maxHit As Integer
    minHit As Integer
    Puntos As Integer
End Type

Type PremiosList
    name As String
    grhIndex As Integer
    Puntos As Integer
End Type

Type NpcInv
    objIndex As Integer
    name As String
    grhIndex As Integer
    amount As Integer
    valor As Long
    objType As Integer
    def As Integer
    maxHit As Integer
    minHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public Type tAccountChar
    sName As String
    iShield As Integer
    iBody As Integer
    iHelm As Integer
    iHead As Integer
    iWeapon As Integer
End Type

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpcInv
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserSelectedChar As Integer
Public UserChars(1 To 8) As tAccountChar
Public CharCount As Byte
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer
Public UserGLD As Long
Public UserPuntos As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserAvisado As Boolean
Public UserMontado As Boolean
Public UserHogar As eCiudad
Public UserFuerza As Byte
Public UserAgilidad As Byte
Public UserMaxHit As Integer
Public UserMinHit As Integer
Public UserEscuMaxDef As Integer
Public UserEscuMinDef As Integer
Public UserHelmetMaxDef As Integer
Public UserHelmetMinDef As Integer
Public UserArmourMaxDef As Integer
Public UserArmourMinDef As Integer

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 4
Public Const NUMSKILLS As Byte = 21
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 16
Public Const NUMRAZAS As Byte = 6

Public userSkills(1 To NUMSKILLS) As Byte
Public userSkillAsign(1 To NUMSKILLS) As Byte
Public skillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean

Public UsingSkill As Integer

Public MD5HushYo As String * 32

Public pingTime As Long

Public Enum E_MODO_LOGIN
    CharLogin = 1
    NewChar = 2
    AccountLogin = 3
End Enum
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

'Used to know if its needed to render inventory.
Public Render_Inventory As Boolean

Public IPdelServidor As String
Public PuertoDelServidor As String


'
'********** FUNCIONES API ***********
'
Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


