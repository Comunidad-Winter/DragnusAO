Attribute VB_Name = "modConsole"
Private Type tLine
    Text As String
    color As Long
    font As Byte
End Type

Private Const MAX_LENGTH As Integer = 255
Private Const MAX_LINES As Integer = 255
Private Const MAX_HEIGHT As Integer = 300
Private Const MAX_WIDTH As Integer = 300


Private lines(0 To MAX_LINES) As tLine
Private lastLine As Integer

Public Sub consoleAdd(ByVal Text As String, ByVal color As Long)
    Dim newLines() As String
    Dim i As Byte
    
    If Len(Text) >= MAX_LENGTH Then
        formatText newLines(), Text
    Else
        ReDim newLines(1) As String
        newLines(0) = Text
    End If
    
    For i = 0 To UBound(newLines) - 1
        If lastLine = MAX_LINES Then
            consoleClear
        Else
            lastLine = lastLine + 1
        End If
        lines(lastLine).Text = newLines(i)
        lines(lastLine).color = color
        lines(lastLine).font = 2
    Next i
End Sub

Private Sub consoleClear()
    Dim i As Byte
    lastLine = 100
    
    For i = 0 To 100
        lines(i) = lines(155 + i)
    Next i
End Sub

Private Sub formatText(newLines() As String, ByVal Text As String)
    Dim i As Integer
    
    Do While Len(Text) > MAX_LENGTH
        ReDim Preserve newLines(0 To i) As String
        newLines(i) = left$(Text, MAX_LENGTH)
        Text = Right(Text, Len(Text) - MAX_LENGTH)
        i = i + 1
    Loop
    ReDim Preserve newLines(0 To i) As String
    newLines(i) = Text

End Sub

Public Sub consoleRender()
    Dim i As Integer
    Dim lineCount As Integer
    
    If lastLine >= 4 Then
        lineCount = lastLine - 4
    Else
        lineCount = 0
    End If
    
    For i = lineCount To lastLine
        dxTextRender 2, lines(i).Text, 530, 10 + (i - lineCount) * 20, lines(i).color
    Next i
End Sub


