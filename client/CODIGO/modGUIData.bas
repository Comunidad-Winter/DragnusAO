Attribute VB_Name = "modGUIDeclares"
Option Explicit

Public Type tRectangle
    x As Integer
    y As Integer
    width As Integer
    height As Integer
End Type

Public Function rectMouseOver(rectangle As tRectangle, mouseX As Integer, mouseY As Integer) As Boolean
    With rectangle
        If mouseX > .x And mouseX < .x + .width Then
            If mouseY > .y And mouseY < .y + .height Then
                rectMouseOver = True
            End If
        End If
    End With
End Function
