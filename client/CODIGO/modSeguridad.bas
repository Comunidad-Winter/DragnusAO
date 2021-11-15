Attribute VB_Name = "modSeguridad"
Public CheatList() As String
Public NumCheats As Integer
Public cheat As String

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
'Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)

'Public Function MD5String(ByVal p As String) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
' compute MD5 digest on a given string, returning the result
'    Dim r As String * 32, T As Long
'    r = Space$(32)
'    T = Len(p)
'    MDStringFix p, T, r
'    MD5String = r
'End Function

Public Function MD5File(ByVal f As String) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space$(32)
    MDFile f, r
    MD5File = r
End Function

Public Sub LoadCheats()

    NumCheats = 33
    ReDim CheatList(1 To NumCheats) As String
    
    CheatList(1) = "!xSpeed.net"
    CheatList(2) = "A Speeder"
    CheatList(3) = "Ao Fast"
    CheatList(4) = "Ao Life"
    CheatList(5) = "AO-BOT"
    CheatList(6) = "AO-Ice"
    CheatList(7) = "AO-ZimX"
    CheatList(8) = "AoMacro"
    CheatList(9) = "ArgenTrap"
    CheatList(10) = "Argentum Pesca"
    CheatList(11) = "Argentum-Pesca"
    CheatList(12) = "Alkon Aoh"
    CheatList(13) = "ANuByS Radar"
    CheatList(14) = "AOItems"
    CheatList(15) = "AOFlechas"
    CheatList(16) = "AoH"
    CheatList(17) = "AoT"
    CheatList(18) = "Chit"
    CheatList(19) = "Easy AO"
    CheatList(20) = "PegaRapido"
    CheatList(21) = "ReymiX Engine"
    CheatList(22) = "cheat engine"
    CheatList(23) = "Cheat.Engine"
    CheatList(24) = "Cheat Engine 5.4"
    CheatList(25) = "Cheat Engine 5.0"
    CheatList(26) = "Cheat Engine 5.1.1"
    CheatList(27) = "Cheat Engine 5.2"
    CheatList(28) = "Cheat Engine 5.3"
    CheatList(29) = "UltraCheat"
    CheatList(30) = "Rlz,Turbinas"
    CheatList(31) = "AOMacro"
    CheatList(32) = "Norton Antivirus"
    CheatList(33) = "Msmsg"
End Sub

Public Function CheckProcesos() As String

Dim i As Long
Dim Proc As PROCESSENTRY32
Dim Snap As Long
Dim ExeName As String

Snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
Proc.dwSize = Len(Proc)
theloop = ProcessFirst(Snap, Proc)
CheckProcesos = ""
Do While theloop <> 0
    ExeName = Proc.szExeFile
    For i = 1 To NumCheats
        If UCase$(Left(ExeName, Len(CheatList(i) & ".exe"))) = UCase$(CheatList(i) & ".exe") Then
            CheckProcesos = Proc.szExeFile
            Exit For
            Exit Do
        End If
        DoEvents
    Next i
    DoEvents
    theloop = ProcessNext(Snap, Proc)
Loop

CloseHandle Snap
End Function

Public Function IsCheating() As Boolean
    If NumCheats > 0 And UBound(CheatList) > 0 Then
        cheat = CheckProcesos
        'Encontro un cheat?
        If cheat <> "" Then
            IsCheating = True
            MsgBox ("Se ha cerrado el juego debido al posible uso de cheats, reloguee.")
        End If
    End If
End Function
