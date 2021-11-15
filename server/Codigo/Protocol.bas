Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue


Private Enum ServerPacketID
    AccountLogged
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    UpdateNeeded            ' REAU
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    userHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMidi                ' TM
    PlayWave                ' TW
    guildList               ' GL
    PlayFireSound           ' FO
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    
    'SAAO
    Premios
    MontuT
    GetProcesos
    SubastaOk
    ShowTorneoForm
    UpdateStrengthAgility
    UpdateArmor
    UpdateEscu
    UpdateCasco
    UpdateHit
    BlacksmithShields
    BlacksmithHelmets
End Enum

Private Enum ClientPacketID
    LoginAccount
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    TALK                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    castspell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequest            'CLANES
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/PASSWD
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    Ping                    '/PING
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    Requestcharinvent    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    RequestTCPStats         '/TCPESSTATS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    EDuelo                  '/DUELO By Blizzard
    SDuelo                  '/SALIRDUELO By Blizzard
    TorneoP                 '/Torneop parejas Blizzard
    Particiar               '/PARTICIPAR blizzard
    NoParticipar            '/NOPARTICIPAR blizzard
    Quest
    PidePremios
    RPremios
    RebornClientPacket
    
    CheckSlot               '/SLOT
End Enum

Private Enum MiClientPacketID
    Montar
    'DeMontar
    Acepto
    NoAcepto
    Subastar
    Ofertar
    ChangeInventorySlotDD
    ShowTorneo
    UserCheat
    Procesos
    SendProcesos
    InfoSub
    SubastaInit
    GoCastle
    Descalificar
    Winner
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_DUELO
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

''
' Handles incoming data.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
'On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(userIndex).incomingData.PeekByte()
    
    If Not (packetID = ClientPacketID.LoginAccount) Then
        'Is the Account actually logged?
        If Not UserList(userIndex).UserAccount.Logged = True Then
            Call closeConnection(userIndex)
            Exit Sub
        Else
            'Does the packet requires a logged user??
            If Not (packetID = ClientPacketID.ThrowDices _
              Or packetID = ClientPacketID.LoginExistingChar _
              Or packetID = ClientPacketID.LoginNewChar) Then
                
                'Is the user actually logged?
                If Not UserList(userIndex).flags.UserLogged Then
                    Call closeConnection(userIndex)
                    Exit Sub
                'He is logged. Reset idle counter if id is valid.
                Else
                    If packetID <= ClientPacketID.CheckSlot Then _
                        UserList(userIndex).Counters.IdleCount = 0
                End If
            ElseIf packetID <= ClientPacketID.CheckSlot Then
                UserList(userIndex).Counters.IdleCount = 0
            End If
        End If
    End If
    
    
    
    Select Case packetID
        Case ClientPacketID.LoginAccount
            Call HandleLoginAccount(userIndex)
        
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(userIndex)
        
        Case ClientPacketID.ThrowDices              'TIRDAD
            Call HandleThrowDices(userIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(userIndex)
        
        Case ClientPacketID.TALK                    ';
            Call HandleTalk(userIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(userIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(userIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(userIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(userIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(userIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(userIndex)
        
        Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
            Call HanldeCombatModeToggle(userIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(userIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(userIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(userIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(userIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(userIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(userIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(userIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(userIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(userIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(userIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(userIndex)
        
        Case ClientPacketID.castspell               'LH
            Call HandleCastSpell(userIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(userIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(userIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(userIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(userIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(userIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(userIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(userIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(userIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(userIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(userIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(userIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(userIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(userIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(userIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(userIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(userIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(userIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(userIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(userIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(userIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(userIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(userIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(userIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(userIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(userIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(userIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(userIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(userIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(userIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(userIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(userIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(userIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(userIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(userIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(userIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(userIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(userIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(userIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(userIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(userIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(userIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(userIndex)
        
        Case ClientPacketID.GuildRequest     'CLANDETAILS
            Call HandleGuildRequest(userIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(userIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(userIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(userIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(userIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(userIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(userIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(userIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(userIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(userIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(userIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(userIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(userIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(userIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(userIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(userIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(userIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(userIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(userIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(userIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(userIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(userIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(userIndex)

        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(userIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(userIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(userIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(userIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(userIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(userIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(userIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(userIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(userIndex)
        
        Case ClientPacketID.ChangePassword          '/PASSWD
            Call HandleChangePassword(userIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(userIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(userIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(userIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(userIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(userIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(userIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(userIndex)
        
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(userIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(userIndex)
        
    
        'GM messages
        Case ClientPacketID.GMMessage               '/GMSG
            Call HandleGMMessage(userIndex)
        
        Case ClientPacketID.showName                '/SHOWNAME
            Call HandleShowName(userIndex)
        
        Case ClientPacketID.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(userIndex)
        
        Case ClientPacketID.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(userIndex)
        
        Case ClientPacketID.GoNearby                '/IRCERCA
            Call HandleGoNearby(userIndex)
        
        Case ClientPacketID.comment                 '/REM
            Call HandleComment(userIndex)
        
        Case ClientPacketID.serverTime              '/HORA
            Call HandleServerTime(userIndex)
        
        Case ClientPacketID.Where                   '/DONDE
            Call HandleWhere(userIndex)
        
        Case ClientPacketID.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(userIndex)
        
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(userIndex)
        
        Case ClientPacketID.WarpChar                '/TELEP
            Call HandleWarpChar(userIndex)
        
        Case ClientPacketID.Silence                 '/SILENCIAR
            Call HandleSilence(userIndex)
        
        Case ClientPacketID.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(userIndex)
        
        Case ClientPacketID.SOSRemove               'SOSDONE
            Call HandleSOSRemove(userIndex)
        
        Case ClientPacketID.GoToChar                '/IRA
            Call HandleGoToChar(userIndex)
        
        Case ClientPacketID.invisible               '/INVISIBLE
            Call HandleInvisible(userIndex)
        
        Case ClientPacketID.GMPanel                 '/PANELGM
            Call HandleGMPanel(userIndex)
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(userIndex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(userIndex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(userIndex)
        
        Case ClientPacketID.Jail                    '/CARCEL
            Call HandleJail(userIndex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(userIndex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(userIndex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(userIndex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(userIndex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(userIndex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(userIndex)
            
        Case ClientPacketID.Requestcharinvent    '/INV
            Call HandleRequestcharinvent(userIndex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(userIndex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(userIndex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(userIndex)
        
        Case ClientPacketID.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(userIndex)
        
        Case ClientPacketID.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(userIndex)
        
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(userIndex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(userIndex)
            
        Case ClientPacketID.BanChar                 '/BAN
            Call HandleBanChar(userIndex)
            
        Case ClientPacketID.UnbanChar               '/UNBAN
            Call HandleUnbanChar(userIndex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(userIndex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(userIndex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(userIndex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(userIndex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(userIndex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(userIndex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(userIndex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(userIndex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(userIndex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(userIndex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(userIndex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(userIndex)
            
        Case ClientPacketID.RainToggle              '/LLUVIA
            Call HandleRainToggle(userIndex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(userIndex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(userIndex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(userIndex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(userIndex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(userIndex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(userIndex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(userIndex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(userIndex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(userIndex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(userIndex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(userIndex)
            
        Case ClientPacketID.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(userIndex)
            
        Case ClientPacketID.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(userIndex)
            
        Case ClientPacketID.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(userIndex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(userIndex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(userIndex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(userIndex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(userIndex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(userIndex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(userIndex)
        
        Case ClientPacketID.BanIP                   '/BANIP
            Call HandleBanIP(userIndex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(userIndex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(userIndex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(userIndex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(userIndex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(userIndex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(userIndex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(userIndex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(userIndex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(userIndex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(userIndex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(userIndex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(userIndex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(userIndex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(userIndex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(userIndex)
        
        Case ClientPacketID.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(userIndex)
        
        Case ClientPacketID.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(userIndex)
        
        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(userIndex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(userIndex)
        
        Case ClientPacketID.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(userIndex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(userIndex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(userIndex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(userIndex)
        
        Case ClientPacketID.AlterPassword           '/APASS
            Call HandleAlterPassword(userIndex)
        
        Case ClientPacketID.AlterMail               '/AEMAIL
            Call HandleAlterMail(userIndex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(userIndex)
        
        Case ClientPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(userIndex)
        
        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(userIndex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(userIndex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(userIndex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(userIndex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(userIndex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(userIndex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(userIndex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(userIndex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(userIndex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(userIndex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(userIndex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(userIndex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(userIndex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(userIndex)
        
        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(userIndex)
        
        Case ClientPacketID.RequestTCPStats         '/TCPESSTATS
            Call HandleRequestTCPStats(userIndex)
        
        Case ClientPacketID.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(userIndex)
        
        Case ClientPacketID.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(userIndex)
        
        Case ClientPacketID.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(userIndex)
        
        Case ClientPacketID.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(userIndex)
        
        Case ClientPacketID.Restart                 '/REINICIAR
            Call HandleRestart(userIndex)
        
        Case ClientPacketID.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(userIndex)
        
        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(userIndex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(userIndex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(userIndex)
            
        Case ClientPacketID.EDuelo
            Call HandleEDuelo(userIndex)
        
        Case ClientPacketID.SDuelo
            Call HandleSDuelo(userIndex)
            
        Case ClientPacketID.TorneoP
            Call HandleTorneo(userIndex)
            
        Case ClientPacketID.Particiar
            Call HandleParticipar(userIndex)
            
        Case ClientPacketID.NoParticipar
            Call HandleNoParticipar(userIndex)
        
        Case ClientPacketID.Quest
            Call HandleQuest(userIndex)
            
        Case ClientPacketID.PidePremios
            Call HandlePremiosRequest(userIndex)
        
        Case ClientPacketID.RPremios
            Call HandleRPremios(userIndex)
            
        Case ClientPacketID.RebornClientPacket
            Call HandleRebornClientPacket(userIndex)
        Case Else
            'ERROR : Abort!
            Call closeConnection(userIndex)
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(userIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(userIndex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(userIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & userIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call closeConnection(userIndex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(userIndex)
    End If
End Sub

Private Sub HandleLoginAccount(ByVal userIndex)
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(userIndex).incomingData)
    Call Buffer.ReadByte
    
    Dim sAccName As String
    Dim sAccPassword As String
    sAccName = Buffer.ReadASCIIString()
    sAccPassword = Buffer.ReadASCIIString()
    If dbCheckAccountData(sAccName, sAccPassword) Then
        If Not isAccountLogged(sAccName) Then
            UserList(userIndex).UserAccount.name = sAccName
            UserList(userIndex).UserAccount.Logged = True
            dbLoadAccountData userIndex
            Call WriteAccountLogged(userIndex)
        Else
            Call WriteErrorMsg(userIndex, "La cuenta ya esta logueada.")
        End If
    Else
        Call WriteErrorMsg(userIndex, "Los datos de la cuenta o la contraseña son incorrectos.")
    End If
    Call UserList(userIndex).incomingData.CopyBuffer(Buffer)
Exit Sub
errhandler:
End Sub



''
' Handles the "LoginExistingChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If UserList(userIndex).incomingData.length < 53 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(userIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim UserName As String
    Dim tempByte As Byte
    
    tempByte = Buffer.ReadByte
    
    Dim version As String
    Dim MD5 As String

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    UserList(userIndex).flags.NoActualizado = Not VersionesActuales(Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger())
    MD5 = Buffer.ReadASCIIString
    
    
    
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(userIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
        Exit Sub
    End If
    
    
    If Not MD5ok(MD5) Then
        Call WriteErrorMsg(userIndex, "Esta version no es la actual, visite www.ao.dveloping.com.ar y checkee las descargas.")
    Else
        If tempByte > 0 And tempByte <= UserList(userIndex).UserAccount.CharCount Then
            UserName = UserList(userIndex).UserAccount.Chars(tempByte)
            If BANCheck(UserName) Then
                Call WriteErrorMsg(userIndex, "Se te ha prohibido la entrada a Argentum debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.com.ar")
            ElseIf Not VersionOK(version) Then
                Call WriteErrorMsg(userIndex, "Esta version del cliente es obsoleta, baja el nuevo cliente desde www.ao.dveloping.com.ar disculpe las molestias.")
            Else
                '¿Ya esta conectado el personaje?
                If CheckForSameName(UserName) Then
                    If UserList(NameIndex(UserName)).Counters.Saliendo Then
                        Call WriteErrorMsg(userIndex, "El usuario está saliendo.")
                    Else
                        Call WriteErrorMsg(userIndex, "Perdon, un usuario con el mismo nombre se há logoeado.")
                    End If
                Else
                    Call ConnectUser(userIndex, UserName)
                End If
            End If
        Else
            Call WriteErrorMsg(userIndex, "Char Invalido.")
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(userIndex).incomingData.CopyBuffer(Buffer)
    
Exit Sub
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    With UserList(userIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = 16 + RandomNumber(0, 2)
        .UserAtributos(eAtributos.Agilidad) = 16 + RandomNumber(0, 2)
        .UserAtributos(eAtributos.Inteligencia) = 16 + RandomNumber(0, 2)
        .UserAtributos(eAtributos.Carisma) = 16 + RandomNumber(0, 2)
        .UserAtributos(eAtributos.Constitucion) = 16 + RandomNumber(0, 2)
    End With
    
    Call WriteDiceRoll(userIndex)
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Debug.Print UserList(userIndex).incomingData.length
    If UserList(userIndex).incomingData.length < 79 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(userIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim skills(NUMSKILLS - 1) As Byte
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    Dim mail As String
    
    Dim MD5 As String
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(userIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(userIndex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(userIndex).ip) Then
        Call WriteErrorMsg(userIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
        
        Exit Sub
    End If
    
    UserName = Buffer.ReadASCIIString()
    
    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    UserList(userIndex).flags.NoActualizado = Not VersionesActuales(Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger(), Buffer.ReadInteger())
    
    race = Buffer.ReadByte()
    gender = Buffer.ReadByte()
    Class = Buffer.ReadByte()
    Call Buffer.ReadBlock(skills, 21)
    
    MD5 = Buffer.ReadASCIIString
    
    If Not MD5ok(MD5) Then
        Call WriteErrorMsg(userIndex, "Esta version del juego no es la correcta, ejecute el AutoUpdater!")
    Else
        If Not VersionOK(version) Then
            Call WriteErrorMsg(userIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            If UserList(userIndex).UserAccount.CharCount >= MAX_ACCOUNT_CHARS Then
                Call WriteErrorMsg(userIndex, "Ya has alcanzado el numero maximo de personajes en esta cuenta.")
            Else
                If dbCharCheck(UserName) Then
                    '¿Existe el personaje?
                     Call WriteErrorMsg(userIndex, "Ya existe el personaje.")
                Else
                    'Tiró los dados antes de llegar acá??
                    If UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
                        Call WriteErrorMsg(userIndex, "Debe tirar los dados antes de poder crear un personaje.")
                    Else
                        Call ConnectNewUser(userIndex, UserName, race, gender, Class, skills)
                    End If
                End If
            End If
        End If
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(userIndex).incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        chat = Buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Dijo: " & chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(userIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, userIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR, 1))
            Else
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor, 1))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Yell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleYell(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        chat = Buffer.ReadASCIIString()
        
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
        Else
            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.name, "Grito: " & chat)
            End If
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(userIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If LenB(chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(chat)
                
                If .flags.Privilegios And PlayerType.User Then
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed, 1))
                Else
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_GM_YELL, 1))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        Dim targetCharIndex As Integer
        Dim targetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = Buffer.ReadInteger()
        chat = Buffer.ReadASCIIString()
        
        targetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        targetPriv = UserList(targetUserIndex).flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            If targetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(userIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                    'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
                    Call WriteConsoleMsg(userIndex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                
                ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                    'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
                    Call WriteConsoleMsg(userIndex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                
                ElseIf Not EstaPCarea(userIndex, targetUserIndex) Then
                    Call WriteConsoleMsg(userIndex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                
                Else
                    '[Consejeros & GMs]
                    If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.name, "Le dijo a '" & UserList(targetUserIndex).name & "' " & chat)
                    End If
                    
                    If LenB(chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(chat)
                        
                        Call WriteChatOverHead(userIndex, chat, .Char.CharIndex, vbBlue)
                        Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, vbBlue)
                        Call FlushBuffer(targetUserIndex)
                        
                        '[CDT 17-02-2004]
                        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                            Call SendData(SendTarget.ToAdminsAreaButConsejeros, userIndex, PrepareMessageChatOverHead("a " & UserList(targetUserIndex).name & "> " & chat, .Char.CharIndex, vbYellow))
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim Heading As eHeading
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Heading = .incomingData.ReadByte()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            If UserList(userIndex).Invent.MonturaObjIndex > 0 Then
                If UserList(userIndex).Clase = eClass.Bandit Then
                    dummy = dummy * (1 + ObjData(UserList(userIndex).Invent.MonturaObjIndex).Speed + 1)
                Else
                    dummy = dummy * (1 + ObjData(UserList(userIndex).Invent.MonturaObjIndex).Speed)
                End If
            End If
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 6100 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .name & " , " & dummy)
                    Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call closeConnection(userIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'salida parche
        If .Counters.Saliendo Then
            Call WriteConsoleMsg(userIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            .Counters.Saliendo = False
            .Counters.Salir = 0
        End If
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(userIndex)
                Call WriteConsoleMsg(userIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(userIndex, Heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(userIndex)
                    Call WriteConsoleMsg(userIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(userIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .Clase <> eClass.Thief Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                'If not under a spell effect, show char
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
        
        If .flags.Muerto = 1 Then
            Call Empollando(userIndex)
        Else
            .flags.EstaEmpo = 0
            .EmpoCont = 0
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(userIndex).incomingData.ReadByte
    
    Call WritePosUpdate(userIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(userIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'Attack
        Dim attackpos As WorldPos
        'Get the attacked position
        attackpos = UserList(userIndex).Pos
        Call HeadtoPos(UserList(userIndex).Char.Heading, attackpos)
        Call userAttacked(userIndex, attackpos, eAttackType.meleeAttack)
        
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(userIndex, "No puedes tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(userIndex)
    End With
End Sub

''
' Handles the "CombatModeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HanldeCombatModeToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.ModoCombate Then
            Call WriteConsoleMsg(userIndex, "Has salido del modo de combate.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, "Has pasado al modo de combate.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        .flags.ModoCombate = Not .flags.ModoCombate
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.seguro Then
            Call WriteSafeModeOff(userIndex)
        Else
            Call WriteSafeModeOn(userIndex)
        End If
        
        .flags.seguro = Not .flags.seguro
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(userIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(userIndex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call WriteAttributes(userIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call WriteSendSkills(userIndex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call WriteMiniStats(userIndex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(userIndex).flags.Comerciando = False
    Call WriteCommerceEnd(userIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = userIndex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user
            Call FlushBuffer(.ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(userIndex)
    End With
End Sub

''
' Handles the "BankEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(userIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(userIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(userIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(userIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim amount As Integer
    Dim ConDrag As Boolean
    Dim X As Integer
    Dim Y As Integer
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        ConDrag = .incomingData.ReadBoolean
        If ConDrag = True Then
            X = .incomingData.ReadInteger
            Y = .incomingData.ReadInteger
        End If
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If amount > 10000 Then Exit Sub 'Don't drop too much gold
            
            If ConDrag Then
                If UserList(userIndex).Pos.X <> X Or UserList(userIndex).Pos.Y <> Y Then
                    Call WriteConsoleMsg(userIndex, "Debes tirar el item bajo tus pies.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            Call TirarOro(amount, userIndex)
            
            Call WriteUpdateGold(userIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
            If ConDrag Then
                If UserList(userIndex).Pos.X <> X Or UserList(userIndex).Pos.Y <> Y Then
                    Call WriteConsoleMsg(userIndex, "Debes tirar el item bajo tus pies.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
                Call DropObj(userIndex, Slot, amount, .Pos.Map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(userIndex, UserList(userIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(userIndex, UserList(userIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWork(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(userIndex, Skill)
            Case Ocultarse
                If .flags.Navegando = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(userIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(userIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(userIndex)
        End Select
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.toadmins, userIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(userIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(userIndex, Slot)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(userIndex, Item)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(userIndex, Item)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim dummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        Dim attackpos As WorldPos
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
                        Or Not InMapBounds(.Pos.Map, X, Y) Then
            Exit Sub
        End If
        
        If Not InRangoVision(userIndex, X, Y) Then
            Call WritePosUpdate(userIndex)
            Exit Sub
        End If
        
        Select Case Skill
            Case eSkill.Proyectiles
                Call LookatTile(userIndex, .Pos.Map, X, Y)
                
                    'Prevent from hitting self
                If UserList(userIndex).flags.TargetUser = userIndex Then
                    Call WriteConsoleMsg(userIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                
                attackpos.Map = UserList(userIndex).Pos.Map
                attackpos.X = X
                attackpos.Y = Y
                                                               
                'Attack!
                Call userAttacked(userIndex, attackpos, eAttackType.rangeAttack)
                '-----------------------------------
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(userIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(userIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                            
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(userIndex, False) Then Exit Sub
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(userIndex) Then
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(userIndex) Then
                        Exit Sub
                    End If
                End If
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call userCastSpell(.flags.Hechizo, userIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(userIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eSkill.Pesca
                dummyInt = .Invent.WeaponEqpObjIndex
                If dummyInt = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(userIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(userIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.Map, X, Y) Then
                    Select Case dummyInt
                        Case CAÑA_PESCA
                            Call DoPescar(userIndex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(userIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(userIndex)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PESCAR))
                Else
                    Call WriteConsoleMsg(userIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(userIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(userIndex, UserList(userIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> userIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(userIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(userIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(userIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(userIndex, "No a quien robarle!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "¡No podés robar en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar
                'Check interval
                If Not IntervaloPermiteTrabajar(userIndex) Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(userIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                dummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If dummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(userIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(userIndex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(dummyInt).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_TALAR))
                        Call DoTalar(userIndex)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(userIndex) Then Exit Sub
                                
                If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(userIndex, .Pos.Map, X, Y)
                
                dummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If dummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    dummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex 'CHECK
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(dummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(userIndex)
                    Else
                        Call WriteConsoleMsg(userIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(userIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(userIndex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(userIndex, tN)
                    Else
                        Call WriteConsoleMsg(userIndex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "No hay ninguna criatura alli!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajar(userIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                            Exit Sub
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                                Call WriteConsoleMsg(userIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(userIndex, "Has sido expulsado por el sistema anti cheats.")
                            Call FlushBuffer(userIndex)
                            Call closeConnection(userIndex)
                            Exit Sub
                        End If
                        
                        Call FundirMineral(userIndex)
                    Else
                        Call WriteConsoleMsg(userIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(userIndex, .Pos.Map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(userIndex)
                        Call EnivarArmadurasConstruibles(userIndex)
                        Call EnivarCascosConstruibles(userIndex)
                        Call EnivarEscudosConstruibles(userIndex)
                        Call WriteShowBlacksmithForm(userIndex)
                    Else
                        Call WriteConsoleMsg(userIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 9 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String
        
        Desc = Buffer.ReadASCIIString()
        GuildName = Buffer.ReadASCIIString()
        site = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(userIndex, Desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.toall, userIndex, PrepareMessageConsoleMsg(.name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            
            'Update tag
             Call RefreshCharStatus(userIndex)
        Else
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(userIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(userIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .Desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(userIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Heading As eHeading
        
        Heading = .incomingData.ReadByte()
        
        'Validate Heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            Call ChangeUserChar(userIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 1 + NUMSKILLS Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call closeConnection(userIndex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
            Call closeConnection(userIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim petIndex As Byte
        
        petIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If petIndex > 0 And petIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(petIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(userIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call NPCVentaItem(userIndex, Slot, amount, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(userIndex, Slot, amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call NPCCompraItem(userIndex, Slot, amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(userIndex, Slot, amount)
    End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim file As String
        Dim title As String
        Dim msg As String
        Dim postFile As String
        
        Dim handle As Integer
        Dim i As Long
        Dim Count As Integer
        
        title = Buffer.ReadASCIIString()
        msg = Buffer.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            file = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
            If FileExist(file, vbNormal) Then
                Count = val(GetVar(file, "INFO", "CantMSG"))
                
                'If there are too many messages, delete the forum
                If Count > MAX_MENSAJES_FORO Then
                    For i = 1 To Count
                        Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
                    Next i
                    Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
                    Count = 0
                End If
            Else
                'Starting the forum....
                Count = 0
            End If
            
            handle = FreeFile()
            postFile = Left$(file, Len(file) - 4) & CStr(Count + 1) & ".for"
            
            'Create file
            Open postFile For Output As handle
            Print #handle, title
            Print #handle, msg
            Close #handle
            
            'Update post count
            Call WriteVar(file, "INFO", "CantMSG", Count + 1)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(userIndex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Desc As String
        Dim codex() As String
        
        Desc = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 6 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        
        Slot = .incomingData.ReadByte()
        amount = .incomingData.ReadLong()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((Slot < 1 Or Slot > MAX_INVENTORY_SLOTS) And Slot <> FLAGORO) _
                        Or amount <= 0 Then Exit Sub
        
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
        'Is the commerce attempt valid??
        If UserList(tUser).ComUsu.DestUsu <> userIndex Then
            Call FinComerciarUsu(userIndex)
            Exit Sub
        End If
        
        'Is he still logged??
        If Not UserList(tUser).flags.UserLogged Then
            Call FinComerciarUsu(userIndex)
            Exit Sub
        Else
            'Is he alive??
            If UserList(tUser).flags.Muerto = 1 Then
                Call FinComerciarUsu(userIndex)
                Exit Sub
            End If
            
            'Has he got enough??
            If Slot = FLAGORO Then
                'gold
                If amount > .Stats.GLD Then
                    Call WriteConsoleMsg(userIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            Else
                'inventory
                If amount > .Invent.Object(Slot).amount Then
                    Call WriteConsoleMsg(userIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            'Prevent offer changes (otherwise people would ripp off other players)
            If .ComUsu.Objeto > 0 Then
                Call WriteConsoleMsg(userIndex, "No puedes cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteConsoleMsg(userIndex, "No podés vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            If .flags.Montado = 1 Then
                If .Invent.MonturaSlot = Slot Then
                    Call WriteConsoleMsg(userIndex, "No podés vender tu montura mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            .ComUsu.Objeto = Slot
            .ComUsu.cant = amount
            
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto = True Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(tUser, .name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)
            End If
            
            Call EnviarObjetoTransaccion(tUser)
        End If
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(userIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(userIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(userIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(userIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(userIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(userIndex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(userIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(userIndex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(userIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(userIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(userIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(userIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = Buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(userIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(userIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(userIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(userIndex, r_ListaDePropuestas(userIndex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(userIndex, r_ListaDePropuestas(userIndex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = Buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(userIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(userIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(userIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(userIndex, UserName, reason, errorStr) Then
            Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, reason)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(userIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(userIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(userIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(userIndex, Buffer.ReadASCIIString())
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim error As String
        
        If Not modGuilds.v_AbrirElecciones(userIndex, error) Then
            Call WriteConsoleMsg(userIndex, error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = Buffer.ReadASCIIString()
        application = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(userIndex, guild, application, errorStr) Then
           Call WriteConsoleMsg(userIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(userIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequest(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    
On Error GoTo errhandler
    With UserList(userIndex)
        
        .incomingData.ReadByte
        
        If .GuildIndex > 0 Then
            If .name = modGuilds.GuildLeader(.GuildIndex) Then
                Call modGuilds.SendGuildLeaderInfo(userIndex)
            Else
                Call modGuilds.SendGuildDetails(userIndex, modGuilds.GuildName(.GuildIndex))
            End If
        Else
            WriteGuildList userIndex, PrepareGuildsList()
        End If
        
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
End Sub

''
' Handles the "Online" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Count As Long
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To LastUser
            If LenB(UserList(i).name) <> 0 Then
                    Count = Count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(userIndex, "Número de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim tUser As Integer
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = userIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(userIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(userIndex)
        End If
        
        Call Cerrar_Usuario(userIndex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(userIndex, .name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(userIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(userIndex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(userIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(userIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> userIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, userIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> userIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, userIndex)
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(userIndex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRest(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(userIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(userIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(userIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(userIndex)
                Call WriteConsoleMsg(userIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(userIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Meditate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Solo podés usar meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             Call WriteConsoleMsg(userIndex, "Sólo las Clases mágicas conocen el arte de la meditación", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(userIndex, "Mana restaurado", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(userIndex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(userIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(userIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Call WriteConsoleMsg(userIndex, "Te estás concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzarás a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = LoopAdEternum
            
            'Show proper FX according to level
            If .Stats.ELV < 15 Then
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARCHICO, LoopAdEternum))
                .Char.FX = FXIDs.FXMEDITARCHICO
                
            ElseIf .Stats.ELV < 25 Then
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARMEDIANO, LoopAdEternum))
                .Char.FX = FXIDs.FXMEDITARMEDIANO
                
            ElseIf .Stats.ELV < 30 Then
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARMEDIANOB, LoopAdEternum))
                .Char.FX = FXIDs.FXMEDITARMEDIANOB
                
            ElseIf .Stats.ELV < 45 Then
                'If criminal(UserIndex) Then
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARGRANDECRIMI, LoopAdEternum))
                    '.Char.FX = FXIDs.FXMEDITARGRANDECRIMI
                'Else
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARGRANDECIUDA, LoopAdEternum))
                    '.Char.FX = FXIDs.FXMEDITARGRANDECIUDA
                'End If
            Else
                'If criminal(UserIndex) Then
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARXGRANDECRIMI, LoopAdEternum))
                    '.Char.FX = FXIDs.FXMEDITARXGRANDECRIMI
                'Else
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARXGRANDECIUDA, LoopAdEternum))
                    '.Char.FX = FXIDs.FXMEDITARXGRANDECIUDA
                'End If
            End If
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(userIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(userIndex)
        Call WriteConsoleMsg(userIndex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Heal" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHP = .Stats.MaxHP
        
        Call WriteUpdateHP(userIndex)
        
        Call WriteConsoleMsg(userIndex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(userIndex, userIndex)
End Sub

''
' Handles the "Help" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call SendHelp(userIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(userIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(userIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(userIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarCOmercioNPC(userIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(userIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(userIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = userIndex Then
                Call WriteConsoleMsg(userIndex, "No puedes comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(userIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> userIndex Then
                Call WriteConsoleMsg(userIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name
            .ComUsu.cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(userIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(userIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(userIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(userIndex)
            End If
        Else
            Call WriteConsoleMsg(userIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(userIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(userIndex)
        Else
            Call EnlistarCaos(userIndex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.Alineacion <> e_Alineacion.Real Then
                 Call WriteChatOverHead(userIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(userIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .Faccion.Alineacion <> e_Alineacion.Caos Then
                 Call WriteChatOverHead(userIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(userIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleReward(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.Alineacion <> e_Alineacion.Real Then
                Call WriteChatOverHead(userIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Call RecompensaArmadaReal(userIndex)
            End If
        Else
            If .Faccion.Alineacion <> e_Alineacion.Caos Then
                Call WriteChatOverHead(userIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Call RecompensaCaos(userIndex)
            End If
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Call SendMOTD(userIndex)
End Sub

''
' Handles the "UpTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    UpTimeStr = time & " dias, " & UpTimeStr
    
    Call WriteConsoleMsg(userIndex, "Uptime: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
    
    'Send auto-reset time
    time = IntervaloAutoReiniciar
    
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    UpTimeStr = time & " dias, " & UpTimeStr
    
    Call WriteConsoleMsg(userIndex, "Próximo mantenimiento automático: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub


''
' Handles the "Inquiry" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(userIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (userIndex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        chat = Buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat))
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                'Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub



''
' Handles the "CentinelReport" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(userIndex, .incomingData.ReadInteger())
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(userIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(userIndex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(userIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub


''
' Handles the "CouncilMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        chat = Buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, userIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, userIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim request As String
        
        request = Buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(userIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(userIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.name)
        Else
            Call Ayuda.Quitar(.name)
            Call Ayuda.Push(.name)
            Call WriteConsoleMsg(userIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BugReport" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Dim N As Integer
        
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = Buffer.ReadASCIIString()
        
        N = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim description As String
        
        description = Buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(userIndex, "La descripción tiene caractéres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Desc = Trim$(description)
                Call WriteConsoleMsg(userIndex, "La descripción a cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildVote" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim vote As String
        Dim errorStr As String
        
        vote = Buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(userIndex, vote, errorStr) Then
            Call WriteConsoleMsg(userIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(userIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Punishments" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim name As String
        Dim Count As Integer
        
        name = Buffer.ReadASCIIString()
        
        If LenB(name) <> 0 Then
            If (InStrB(name, "\") <> 0) Then
                name = Replace(name, "\", "")
            End If
            If (InStrB(name, "/") <> 0) Then
                name = Replace(name, "/", "")
            End If
            If (InStrB(name, ":") <> 0) Then
                name = Replace(name, ":", "")
            End If
            If (InStrB(name, "|") <> 0) Then
                name = Replace(name, "|", "")
            End If
            
            If FileExist(CharPath & name & ".chr", vbNormal) Then
                Count = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
                If Count = 0 Then
                    Call WriteConsoleMsg(userIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                Else
                    While Count > 0
                        Call WriteConsoleMsg(userIndex, Count & " - " & GetVar(CharPath & name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                        Count = Count - 1
                    Wend
                End If
            Else
                Call WriteConsoleMsg(userIndex, "Personaje """ & name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo errhandler
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Pass As String
        
        'Get password and validate it if necessary
        Pass = Buffer.ReadASCIIString()
        
        If Len(Pass) < 6 Then
             Call WriteConsoleMsg(userIndex, "El password debe tener al menos 6 caractéres.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteVar(CharPath & UserList(userIndex).name & ".chr", "INIT", "Password", Pass)
            
            'Everything is right, change password
            Call WriteConsoleMsg(userIndex, "El password ha sido cambiado.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
errhandler:

End Sub

''
' Handles the "Gamble" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Integer
        
        amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(userIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount < 1 Then
            Call WriteChatOverHead(userIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount > 5000 Then
            Call WriteChatOverHead(userIndex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < amount Then
            Call WriteChatOverHead(userIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + amount
                Call WriteChatOverHead(userIndex, "Felicidades! Has ganado " & CStr(amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - amount
                Call WriteChatOverHead(userIndex, "Lo siento, has perdido " & CStr(amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(userIndex)
        End If
    End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(userIndex, ConsultaPopular.doVotar(userIndex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        
        amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If amount > 0 And amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - amount
             .Stats.GLD = .Stats.GLD + amount
             Call WriteChatOverHead(userIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             Call WriteChatOverHead(userIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(userIndex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Noble Then
           'Quit the Royal Army?
           If .Faccion.Alineacion = e_Alineacion.Real Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
                   Call ExpulsarFaccionReal(userIndex)
                   Call WriteChatOverHead(userIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(userIndex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
            'Quit the Chaos Legion??
           ElseIf .Faccion.Alineacion = e_Alineacion.Caos Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 1 Then
                   Call ExpulsarFaccionCaos(userIndex)
                   Call WriteChatOverHead(userIndex, "Ya volverás arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(userIndex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
           Else
               Call WriteChatOverHead(userIndex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
           End If
        End If
    End With
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim amount As Long
        
        amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(userIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(userIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If amount > 0 And amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + amount
            .Stats.GLD = .Stats.GLD - amount
            Call WriteChatOverHead(userIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(userIndex)
        Else
            Call WriteChatOverHead(userIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Text As String
        
        Text = Buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg(LCase$(.name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(userIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim error As String
        
        Select Case .Faccion.Alineacion
            Case e_Alineacion.Real
                .FundandoGuildAlineacion = ALINEACION_ARMADA
            Case e_Alineacion.Neutro
                .FundandoGuildAlineacion = ALINEACION_NEUTRO
            Case e_Alineacion.Caos
                .FundandoGuildAlineacion = ALINEACION_LEGION
        End Select
        
        If modGuilds.PuedeFundarUnClan(userIndex, .FundandoGuildAlineacion, error) Then
            Call WriteShowGuildFundationForm(userIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(userIndex, error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub




''
' Handles the "GuildMemberList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(userIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(userIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        
        message = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg(.name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(userIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.Alineacion = e_Alineacion.Real Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(userIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(userIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.Alineacion = e_Alineacion.Caos Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(userIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(userIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        
        UserName = Buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.Map, X, Y).userIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(userIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(userIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Comment" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleComment(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim comment As String
        comment = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(userIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(userIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Dim i As Long
        Dim List1 As String
        Dim list2 As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then
                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        List1 = List1 & Npclist(i).name & "(" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & "), "
                    Else
                        list2 = list2 & Npclist(i).name & "(" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & "), "
                    End If
                End If
            Next i
            
            If LenB(List1) <> 0 Then
                List1 = Left$(List1, Len(List1) - 2)
            Else
                List1 = "No hay NPCS Hostiles"
            End If
            
            If LenB(list2) <> 0 Then
                list2 = Left$(list2, Len(list2) - 2)
            Else
                list2 = "No hay más NPCS"
            End If
            
            Call WriteConsoleMsg(userIndex, "Npcs Hostiles en mapa: " & List1, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(userIndex, "Otros Npcs en mapa: " & list2, FontTypeNames.FONTTYPE_INFO)
            Call LogGM(.name, "Numero enemigos en mapa " & Map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WarpUserChar(userIndex, .flags.TargetMap, .flags.targetX, .flags.targetY, True)
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.targetX & " Y:" & .flags.targetY & " Map:" & .Pos.Map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 7 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim Map As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        Map = Buffer.ReadInteger()
        X = Buffer.ReadByte()
        Y = Buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = userIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(Map, X, Y) Then
                    Call WarpUserChar(tUser, Map, X, Y, True)
                    Call WriteConsoleMsg(userIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(userIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.name, "/silenciar " & UserList(tUser).name)
                
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(userIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(userIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WarpUserChar(userIndex, UserList(tUser).Pos.Map, UserList(tUser).Pos.X, UserList(tUser).Pos.Y + 1, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Invisible" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(userIndex)
        Call LogGM(.name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(userIndex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(userIndex, names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then _
                    users = users & " (*)"
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(userIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(userIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleJail(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 6 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        jailTime = Buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(userIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(userIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(userIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(userIndex, "No podés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & time)
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name)
                        Call LogGM(.name, " encarcelo a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(userIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
        Else
            Call WriteConsoleMsg(userIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(userIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(userIndex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & time)
                        
                        Call WriteConsoleMsg(userIndex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "EditChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/28/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 8 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim commandString As String
        Dim N As Byte
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = userIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = Buffer.ReadByte()
        Arg1 = Buffer.ReadASCIIString()
        Arg2 = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, Body y level
                    valido = tUser = userIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y Body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = userIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = userIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If
        
        If valido Then
            Select Case opcion
                Case eEditOptions.eo_Gold
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) < 5000000 Then
                            UserList(tUser).Stats.GLD = val(Arg1)
                            Call WriteUpdateGold(tUser)
                        Else
                            Call WriteConsoleMsg(userIndex, "No esta permitido utilizar valores mayores. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                
                Case eEditOptions.eo_Experience
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) < 15995001 Then
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        Else
                            Call WriteConsoleMsg(userIndex, "No esta permitido utilizar valores mayores a mucho. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                
                Case eEditOptions.eo_Body
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Body", Arg1)
                        Call WriteConsoleMsg(userIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_Head
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Head", Arg1)
                        Call WriteConsoleMsg(userIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_CriminalsKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXUSERMATADOS Then
                            UserList(tUser).Faccion.CriminalesMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).Faccion.CriminalesMatados = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_CiticensKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXUSERMATADOS Then
                            UserList(tUser).Faccion.CiudadanosMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).Faccion.CiudadanosMatados = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_Level
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(userIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        UserList(tUser).Stats.ELV = val(Arg1)
                    End If
                    
                    Call WriteUpdateUserStats(userIndex)
                
                Case eEditOptions.eo_Class
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        For LoopC = 1 To NUMClaseS
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMClaseS Then
                            Call WriteConsoleMsg(userIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Clase = LoopC
                        End If
                    End If
                
                Case eEditOptions.eo_Skills
                    For LoopC = 1 To NUMSKILLS
                        If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                    Next LoopC
                    
                    If LoopC > NUMSKILLS Then
                        Call WriteConsoleMsg(userIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If tUser <= 0 Then
                            Call WriteVar(CharPath & UserName & ".chr", "Skills", "SK" & LoopC, Arg2)
                            Call WriteConsoleMsg(userIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                        End If
                    End If
                
                Case eEditOptions.eo_SkillPointsLeft
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "STATS", "SkillPtsLibres", Arg1)
                        Call WriteConsoleMsg(userIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(tUser).Stats.SkillPts = val(Arg1)
                    End If
                
                Case eEditOptions.eo_Sex
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Arg1 = UCase$(Arg1)
                        If (Arg1 = "MUJER") Then
                            UserList(tUser).Genero = eGenero.Mujer
                        ElseIf (Arg1 = "HOMBRE") Then
                            UserList(tUser).Genero = eGenero.Hombre
                        End If
                    End If
                
                Case eEditOptions.eo_Raza
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(userIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Arg1 = UCase$(Arg1)
                        If (Arg1 = "HUMANO") Then
                            UserList(tUser).Raza = eRaza.Humano
                        ElseIf (Arg1 = "ELFO") Then
                            UserList(tUser).Raza = eRaza.Elfo
                        ElseIf (Arg1 = "DROW") Then
                            UserList(tUser).Raza = eRaza.ElfoOscuro
                        ElseIf (Arg1 = "ENANO") Then
                            UserList(tUser).Raza = eRaza.Enano
                        ElseIf (Arg1 = "GNOMO") Then
                            UserList(tUser).Raza = eRaza.Gnomo
                        ElseIf (Arg1 = "ORCO") Then
                            UserList(tUser).Raza = eRaza.Orco
                        End If
                    End If
                
                Case Else
                    Call WriteConsoleMsg(userIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
            End Select
        End If
        
        'Log it!
        commandString = "/MOD "
        
        Select Case opcion
            Case eEditOptions.eo_Gold
                commandString = commandString & "ORO "
            
            Case eEditOptions.eo_Experience
                commandString = commandString & "EXP "
            
            Case eEditOptions.eo_Body
                commandString = commandString & "Body "
            
            Case eEditOptions.eo_Head
                commandString = commandString & "HEAD "
            
            Case eEditOptions.eo_CriminalsKilled
                commandString = commandString & "CRI "
            
            Case eEditOptions.eo_CiticensKilled
                commandString = commandString & "CIU "
            
            Case eEditOptions.eo_Level
                commandString = commandString & "LEVEL "
            
            Case eEditOptions.eo_Class
                commandString = commandString & "Clase "
            
            Case eEditOptions.eo_Skills
                commandString = commandString & "SKILLS "
            
            Case eEditOptions.eo_SkillPointsLeft
                commandString = commandString & "SKILLSLIBRES "
                
            Case eEditOptions.eo_Nobleza
                commandString = commandString & "NOB "
                
            Case eEditOptions.eo_Asesino
                commandString = commandString & "ASE "
                
            Case eEditOptions.eo_Sex
                commandString = commandString & "SEX "
                
            Case eEditOptions.eo_Raza
                commandString = commandString & "Raza "
                
            Case Else
                commandString = commandString & "UNKOWN "
        End Select
        
        commandString = commandString & Arg1 & " " & Arg2
        
        If valido Then _
            Call LogGM(.name, commandString & " " & UserList(tUser).name)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal userIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
                
        Dim targetName As String
        Dim targetIndex As Integer
        
        targetName = Replace$(Buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If targetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(userIndex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(userIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(userIndex, targetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(userIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(userIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(userIndex, UserName)
            Else
                Call WriteConsoleMsg(userIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Requestcharinvent" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestcharinvent(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(userIndex, UserName)
            Else
                Call SendUserInvTxt(userIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(userIndex, UserName)
            Else
                Call SendUserBovedaTxt(userIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Requestcharskills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(userIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(userIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = userIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    .flags.Muerto = 0
                    .Stats.MinHP = .Stats.MaxHP
                    
                    Call DarCuerpoDesnudo(tUser)
                    
                    Call ChangeUserChar(tUser, .Char.body, .OrigChar.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call WriteConsoleMsg(tUser, .name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal userIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(userIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.Map = .Pos.Map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(userIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub



''
' Handles the "Kick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKick(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(userIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(.name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call closeConnection(tUser)
                    Call LogGM(.name, "Echo a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(userIndex, "Estás loco?? como vas a piñatear un gm!!!! :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "No está online", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(userIndex, UserName, reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(userIndex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(userIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(userIndex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .name & " te há trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y + 1, True)
                    Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(userIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(userIndex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.name, "/RESETINV " & Npclist(.flags.TargetNPC).name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & message)
                message = .name & ": " & message
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(userIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(userIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(userIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = Buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(userIndex, "Clan " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(userIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "/CT " & mapa & "," & X & "," & Y)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(userIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(userIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(.Pos.Map, ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.targetX
        Y = .flags.targetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(userIndex).name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(mapa, .ObjInfo.amount, mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(.TileExit.Map, 1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.toall, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim tUser As Integer
        Dim Desc As String
        
        Desc = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(userIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.Map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 6 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToNeutralesYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub


''
' Handles the "TalkAsNPC" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(userIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(.Pos.Map, MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(userIndex).name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legión Oscura.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim lista As String
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(userIndex, "(" & X & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(userIndex, "Usuario offline, Echando de los consejos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(userIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de la Legión Oscura", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legión Oscura", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.name, tLog)
            Call WriteConsoleMsg(userIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(userIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(userIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(userIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(.name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call closeConnection(tIndex)
                    End If
                    
                    'ponemos el flag de ban a 1
                    Call dbWriteInteger("charinit", "Ban", member, 1)
                    
                    'ponemos la pena
                    'Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    'Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    'Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
                Next LoopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "BanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 6 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If Buffer.ReadBoolean() Then
            bannedIP = Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte()
        Else
            tUser = NameIndex(Buffer.ReadASCIIString())
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(userIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedIP = UserList(tUser).ip
            End If
        End If
        
        reason = Buffer.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(userIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg(.name & " baneó la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser
                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call BanCharacter(userIndex, UserList(i).name, "IP POR " & reason)
                        End If
                    End If
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(userIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        tObj = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.name, "/CI: " & tObj)
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(userIndex, "ATENCION: FUERON CREADOS ***100*** ITEMS!, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.amount = 100
        Objeto.ObjIndex = tObj
        Call MakeObj(.Pos.Map, Objeto, .Pos.Map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.name, "/DEST")
        
        If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(userIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(.Pos.Map, 10000, .Pos.Map, .Pos.X, .Pos.Y)
    End With
End Sub
''
' Handles the "ForceMIDIAll" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.toall, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.toall, 0, PrepareMessagePlayWave(waveID))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(userIndex).incomingData.length < 6 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = Buffer.ReadASCIIString()
        punishment = Buffer.ReadByte
        NewText = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(userIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.name) & ": <" & NewText & "> " & Date & " " & time)
                    
                    Call WriteConsoleMsg(userIndex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal userIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(userIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(userIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(userIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = Buffer.ReadASCIIString() 'Que UserName?
        Slot = Buffer.ReadByte() 'Que Slot?
        tIndex = NameIndex(UserName)  'Que user index?
        
        Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & UserName)
           
        If tIndex > 0 Then
            If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(userIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(userIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
            End If
        Else
            Call WriteConsoleMsg(userIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
    With UserList(userIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.name) <> "MARAXUS" Then Exit Sub
        
        Call WriteConsoleMsg(userIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Restart" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.name) <> "MARAXUS" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.name, .name & " reinicio el mundo")
        
        Call ReiniciarServidor(True)
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado a los objetos.")
        
        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha recargado los INITs.")
        
        Call LoadSini
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.name, .name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(userIndex, "Npcs.dat y npcsHostiles.dat recargados.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "RequestTCPStats" message
' @param UserIndex The index of the user sending the message

Public Sub HandleRequestTCPStats(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Send the TCP`s stadistics
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
                
        Dim list As String
        Dim Count As Long
        Dim i As Long
        
        Call LogGM(.name, .name & " ha pedido las estadisticas del TCP.")
    
        Call WriteConsoleMsg(userIndex, "Los datos están en BYTES.", FontTypeNames.FONTTYPE_INFO)
        
        'Send the stats
        With TCPESStats
            Call WriteConsoleMsg(userIndex, "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(userIndex, "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(userIndex, "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
        End With
        
        'Search for users that are working
        For i = 1 To LastUser
            With UserList(i)
                If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
                    If .outgoingData.length > 0 Then
                        list = list & .name & " (" & CStr(.outgoingData.length) & "), "
                        Count = Count + 1
                    End If
                End If
            End With
        Next i
        
        Call WriteConsoleMsg(userIndex, "Posibles pjs trabados: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(userIndex, list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub



''
' Handle the "ShowServerForm" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha borrado los SOS")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado todos los chars")

        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la información sobre el BackUp")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informacion sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS".
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim tStr As String
    
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(userIndex).Pos.Map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(userIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            MapInfo(UserList(userIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            MapInfo(UserList(userIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(userIndex).incomingData.length < 2 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            MapInfo(UserList(userIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim tStr As String
    
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(userIndex).Pos.Map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(userIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(userIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal userIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim tStr As String
    
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(userIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(userIndex).Pos.Map & ".dat", "Mapa" & UserList(userIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(userIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(userIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(userIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "SaveMap" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(userIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        
        guild = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(userIndex, guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha hecho un backup")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0
        End With
    
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If
        
        If centinelaActivado Then
            Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        newName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(userIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(userIndex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(userIndex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(userIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(userIndex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)
                                
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & UserName)
                            Else
                                Call WriteConsoleMsg(userIndex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "AlterName" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = Buffer.ReadASCIIString()
        newMail = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(userIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(userIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(userIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                    UserList(userIndex).email = newMail
                End If
                
                Call LogGM(.name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "AlterPassword" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(userIndex).incomingData.length < 5 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(userIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(userIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(userIndex, "Password de " & UserName & " cambiado a: " & Password, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        Call LogGM(.name, "Sumoneo con respawn " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)
        
    End With
End Sub

''
' Handle the "ImperialArmour" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim Index As Byte
        Dim ObjIndex As Integer
        
        Index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index
            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "ChaosArmour" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 4 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim Index As Byte
        Dim ObjIndex As Integer
        
        Index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index
            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(userIndex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(userIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(userIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer
    
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "/APAGAR")
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(.name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "ResetFactions" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then _
                Call ResetFacciones(tUser)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(userIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(userIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(userIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(userIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.toall, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "SetMOTD" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(userIndex).incomingData.length < 3 Then
        Err.raise UserList(userIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    With UserList(userIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = Buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(userIndex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(userIndex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandlePing(ByVal userIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(userIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(userIndex)
    End With
End Sub

Public Sub WriteAccountLogged(ByVal userIndex As Integer)
    Dim iShield As Integer
    Dim iHead As Integer
    Dim iHelm As Integer
    Dim iBody As Integer
    Dim iWeapon As Integer

    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.AccountLogged)
        Call .WriteByte(UserList(userIndex).UserAccount.CharCount)
        Dim i As Byte
        For i = 1 To UserList(userIndex).UserAccount.CharCount
            Call dbGetAccountCharInfo(UserList(userIndex).UserAccount.Chars(i), iBody, iHead, iWeapon, iShield, iHelm)
            Call .WriteASCIIString(UserList(userIndex).UserAccount.Chars(i))
            Call .WriteInteger(iBody)
            Call .WriteInteger(iHead)
            Call .WriteInteger(iWeapon)
            Call .WriteInteger(iShield)
            Call .WriteInteger(iHelm)
        Next i
    End With
End Sub



''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.Logged)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal userIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.NPCSwing)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCKillUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.UserSwing)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateNeeded" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateNeeded(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateNeeded" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.UpdateNeeded)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOn" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOff" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NobilityLost" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.NobilityLost)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(userIndex).Stats.MinSta)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(userIndex).Stats.MinMAN)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(userIndex).Stats.MinHP)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(userIndex).Stats.GLD)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(userIndex).Stats.Exp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal userIndex As Integer, ByVal Map As Integer, ByVal version As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(version)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(userIndex).Pos.X)
        Call .WriteByte(UserList(userIndex).Pos.Y)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the Body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal userIndex As Integer, ByVal target As PartesCuerpo, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCHitUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(target)
        Call .WriteInteger(damage)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal userIndex As Integer, ByVal damage As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.userHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal userIndex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the Body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal userIndex As Integer, ByVal target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(target)
        Call .WriteInteger(damage)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the Body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal userIndex As Integer, ByVal target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(target)
        Call .WriteInteger(damage)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal userIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, color))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal userIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal userIndex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal userIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(userIndex)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(userIndex).Char.CharIndex)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Body Body index of the new character.
' @param    head Head index of the new character.
' @param    Heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal userIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte, ByVal Aura As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, name, criminal, privileges, Aura))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal userIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal userIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Body Body index of the new character.
' @param    head Head index of the new character.
' @param    Heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal userIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    'Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(Body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal userIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal userIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal userIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal userIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal userIndex As Integer, ByVal wave As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayWave" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal userIndex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim Tmp As String
    Dim i As Long
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayFireSound" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayFireSound(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayFireSound" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayFireSound())
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(userIndex).Pos.X)
        Call .WriteByte(UserList(userIndex).Pos.Y)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal userIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(userIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(userIndex).Stats.MinHP)
        Call .WriteInteger(UserList(userIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(userIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(userIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(userIndex).Stats.MinSta)
        Call .WriteLong(UserList(userIndex).Stats.GLD)
        Call .WriteByte(UserList(userIndex).Stats.ELV)
        Call .WriteLong(UserList(userIndex).Stats.ELU)
        Call .WriteLong(UserList(userIndex).Stats.Exp)
        Call .WriteLong(UserList(userIndex).Stats.puntos)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal userIndex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal userIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(UserList(userIndex).Invent.Object(Slot).amount)
        Call .WriteBoolean(UserList(userIndex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.def)
        Call .WriteLong(obData.Valor \ REDUCTOR_PRECIOVENTA)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal userIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(userIndex).UserAccount.Boveda.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteASCIIString(obData.name)
        Call .WriteInteger(UserList(userIndex).UserAccount.Boveda.Object(Slot).amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.def)
        Call .WriteLong(obData.Valor)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal userIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(userIndex).Stats.UserHechizos(Slot))
        
        If UserList(userIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(userIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub



Public Sub WriteBlacksmithShields(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(EscudosHerrero()))
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithShields)
        
        For i = 1 To UBound(EscudosHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(EscudosHerrero(i)).SkHerreria <= UserList(userIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(EscudosHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(EscudosHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithHelmets(ByVal userIndex As Integer)
'***************************************************
'Author: Agustin Andreucci (Blizzard)
'Last Modification: 05/17/06
'Writes the "BlacksmithHelmets" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(CascosHerrero()))
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithHelmets)
        
        For i = 1 To UBound(CascosHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(CascosHerrero(i)).SkHerreria <= UserList(userIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(CascosHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(CascosHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub



''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(userIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(userIndex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(userIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(userIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.name)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal userIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal userIndex As Integer, ByVal ObjIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    obj The object to be set in the NPC's inventory window.
' @param    price The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal userIndex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.name)
        Call .WriteInteger(Obj.amount)
        Call .WriteLong(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.def)
    End With
Exit Sub



errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(userIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(userIndex).Stats.MinAGU)
        Call .WriteByte(UserList(userIndex).Stats.MaxHam)
        Call .WriteByte(UserList(userIndex).Stats.MinHam)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub



''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(userIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(userIndex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(userIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(userIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(userIndex).Clase)
        Call .WriteLong(UserList(userIndex).Counters.Pena)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal userIndex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal userIndex As Integer, ByVal title As String, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(message)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal userIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DiceRoll" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma))
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SendSkills" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(userIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal userIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal userIndex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal userIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal userIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal userIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal userIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal Reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        Call .WriteLong(Reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal userIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String, ByRef peaceRequest() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guilds' list
        For i = LBound(peaceRequest()) To UBound(peaceRequest())
            Tmp = Tmp & peaceRequest(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal userIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByRef codex() As String, ByVal guildDesc As String, ByVal Reputation As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim temp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Dim mOnline() As String
        Dim members() As String
        
        mOnline = Split(modGuilds.m_ListaDeMiembrosOnline(userIndex, UserList(userIndex).GuildIndex), ",")
        Call .WriteInteger(UBound(mOnline))
        For i = 0 To UBound(mOnline)
            Call .WriteASCIIString(mOnline(i))
        Next i
        
        members = Split(modGuilds.m_ListaDeMiembros(userIndex, UserList(userIndex).GuildIndex), ",")
        Call .WriteInteger(memberCount)
        For i = 0 To UBound(members)
            Call .WriteASCIIString(members(i))
        Next i
        
        Call .WriteASCIIString(alignment)
        
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(userIndex) 'Actualizamos la posicion del usuario (AUTO-L)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal userIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal userIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(ObjData(ObjIndex).name)
        Call .WriteLong(amount)
        Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
        Call .WriteInteger(ObjData(ObjIndex).MinHIT)
        Call .WriteInteger(ObjData(ObjIndex).def)
        Call .WriteLong(ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal userIndex As Integer, ByVal night As Byte)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.SendNight)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal userIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal userIndex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal userIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errhandler
    Call UserList(userIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal userIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With UserList(userIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(userIndex, sndData)
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long, Optional ByVal Consola As Byte = 0) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)

        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        Call .WriteByte(Consola)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal effectID As Integer, ByVal lifeTime As Integer, Optional ByVal target As Byte = 0, _
                                        Optional ByVal projectileID As Byte = 0, Optional ByVal mapX As Byte = 0, Optional ByVal mapY As Byte = 0, _
                                        Optional ByVal origX As Byte = 0, Optional ByVal origY As Byte = 0, Optional ByVal onHitEffect As Byte = 0, _
                                        Optional ByVal onHitTarget As Byte = 0) As String
                                            
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        
        Call .WriteByte(effectID)
        Call .WriteByte(projectileID)
        
        Call .WriteByte(target)
        
        Call .WriteInteger(lifeTime)
        
        If target = 0 Then
            Call .WriteInteger(CharIndex)
        Else
            Call .WriteByte(mapX)
            Call .WriteByte(mapY)
        End If
        
        If projectileID > 0 Then
            Call .WriteByte(origX)
            Call .WriteByte(origY)
            Call .WriteByte(onHitEffect)
            Call .WriteByte(onHitTarget)
        End If
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PlayWave" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PlayFireSound" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayFireSound() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PlayFireSound" and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayFireSound)
        
        PrepareMessagePlayFireSound = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    Body Body index of the new character.
' @param    head Head index of the new character.
' @param    Heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte, ByVal Aura As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(head)
        Call .WriteByte(Heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteInteger(Aura)
        Call .WriteASCIIString(name)
        Call .WriteByte(criminal)
        Call .WriteByte(privileges)
        
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    Body Body index of the new character.
' @param    head Head index of the new character.
' @param    Heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal head As Integer, ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Aura As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(head)
        Call .WriteByte(Heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteInteger(Aura)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal userIndex As Integer, Faccion As Byte, Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(userIndex).Char.CharIndex)
        Call .WriteByte(Faccion)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WritePremios(ByVal userIndex As Integer)
    
    If CantPremios = 0 Then Exit Sub
    
    Dim i As Integer
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.Premios)
        Call .WriteInteger(CantPremios)
        Call .WriteLong(UserList(userIndex).Stats.puntos)
    End With
    
    For i = 1 To CantPremios
        With UserList(userIndex).outgoingData
            Call .WriteASCIIString(ObjData(PremiosInfo(i).ObjIndex).name)
            Call .WriteInteger(PremiosInfo(i).puntos)
            Call .WriteInteger(ObjData(PremiosInfo(i).ObjIndex).GrhIndex)
        End With
    Next i
    
End Sub
Public Sub HandlePremiosRequest(ByVal userIndex As Integer)
    UserList(userIndex).incomingData.ReadByte
    
    Call WritePremios(userIndex)
End Sub
Public Sub WriteMontuToggle(userIndex)

With UserList(userIndex).outgoingData
    .WriteByte (ServerPacketID.MontuT)
    If UserList(userIndex).Invent.MonturaObjIndex > 0 Then
        If UserList(userIndex).Clase = eClass.Bandit Then
            .WriteByte (ObjData(UserList(userIndex).Invent.MonturaObjIndex).Speed + 1) 'El bandido es mas rapido.
        Else
            .WriteByte (ObjData(UserList(userIndex).Invent.MonturaObjIndex).Speed)
        End If
    Else
        .WriteByte (0)
    End If
End With

End Sub
Public Sub WriteGetProcesos(ByVal userIndex As Integer, ByVal GetIndex As Integer)
    
    If Not GetIndex > 0 And Not userIndex > 0 Then Exit Sub

    Call UserList(GetIndex).outgoingData.WriteByte(ServerPacketID.GetProcesos)
    Call UserList(GetIndex).outgoingData.WriteInteger(userIndex)
    
End Sub

Public Sub WriteSubastaOk(ByVal userIndex As Integer)
    UserList(userIndex).outgoingData.WriteByte (ServerPacketID.SubastaOk)
End Sub
Public Sub WriteUpdateStrengthAgility(ByVal userIndex As Integer)
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrengthAgility)
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
End Sub
Public Sub WriteUpdateArmor(ByVal userIndex As Integer)
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateArmor)
        If UserList(userIndex).Invent.ArmourEqpObjIndex Then
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MaxDef)
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MinDef)
        Else
            Call .WriteInteger(0)
            Call .WriteInteger(0)
        End If
    End With
End Sub
Public Sub WriteUpdateEscu(ByVal userIndex As Integer)
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateEscu)
        If UserList(userIndex).Invent.EscudoEqpObjIndex Then
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).MaxDef)
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).MinDef)
        Else
            Call .WriteInteger(0)
            Call .WriteInteger(0)
        End If
    End With
End Sub
Public Sub WriteUpdateCasco(ByVal userIndex As Integer)
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateCasco)
        If UserList(userIndex).Invent.CascoEqpObjIndex Then
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).MaxDef)
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).MinDef)
        Else
            Call .WriteInteger(0)
            Call .WriteInteger(0)
        End If
    End With
End Sub
Public Sub WriteUpdateHit(ByVal userIndex As Integer)
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHit)
        If UserList(userIndex).Invent.WeaponEqpObjIndex Then
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).MaxHIT)
            Call .WriteInteger(ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).MinHIT)
        Else
            Call .WriteInteger(0)
            Call .WriteInteger(0)
        End If
    End With
End Sub
Public Sub WriteShowTorneoForm(ByVal userIndex As Integer)

On Error GoTo errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(userIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowTorneoForm)
        
        For i = 1 To ColaTorneo.Longitud
            Tmp = Tmp & ColaTorneo.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
    
Exit Sub

errhandler:
    If Err.Number = UserList(userIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(userIndex)
        Resume
    End If
    
End Sub
Public Sub HandleRebornClientPacket(ByVal userIndex As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
'On Error Resume Next
    Dim packetID As Byte
    
    UserList(userIndex).incomingData.ReadByte
    packetID = UserList(userIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.ThrowDices _
      Or packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar) Then
        
        'Is the user actually logged?
        If Not UserList(userIndex).flags.UserLogged Then
            Call closeConnection(userIndex)
            Exit Sub
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= ClientPacketID.CheckSlot Then
                UserList(userIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= ClientPacketID.CheckSlot Then
        UserList(userIndex).Counters.IdleCount = 0
    End If
    
    
    Select Case packetID
        Case MiClientPacketID.Montar
            Call HandleMontar(userIndex)
        Case MiClientPacketID.Acepto
            Call HandleSubastar(userIndex)
        Case MiClientPacketID.Ofertar
            Call HandleOfertar(userIndex)
        Case MiClientPacketID.ChangeInventorySlotDD
            Call HandleChangeInventorySlot(userIndex)
        Case MiClientPacketID.ShowTorneo
            Call HandleShowTorneo(userIndex)
        Case MiClientPacketID.UserCheat
            Call HandleUserCheat(userIndex)
        Case MiClientPacketID.Procesos
            Call HandleProcesos(userIndex)
        Case MiClientPacketID.SendProcesos
            Call HandleSendProcesos(userIndex)
        Case MiClientPacketID.SubastaInit
            Call HandleSubastaInit(userIndex)
        Case MiClientPacketID.InfoSub
            Call HandleInfoSub(userIndex)
        Case MiClientPacketID.ShowTorneo
            Call HandleShowTorneo(userIndex)
        Case MiClientPacketID.GoCastle
            Call HandleGoCastle(userIndex)
        Case MiClientPacketID.Descalificar
            Call HandleDescalificar(userIndex)
        Case MiClientPacketID.Winner
            Call HandleWinner(userIndex)
    End Select
    
End Sub
