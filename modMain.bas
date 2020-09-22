Attribute VB_Name = "modMain"
Public LastStar(2) As Long
Public Type star
    X As Long
    Y As Long
End Type

Public Type RainDrop
    X As Long
    Y As Long
    Text As String
End Type
Public Stars() As star
Public StarsLow As Long
Public StarsMid As Long
Public StarsHi As Long
Public RainIndex As Long
Public RainingText() As RainDrop
Public LastLetter As Boolean
Public Steps As Long

Public Function StartInit()
    Dim i As Long
    For i = 0 To UBound(Stars) - 1
        Stars(i).X = Rnd * Screen.Width + 1
        Stars(i).Y = Rnd * Screen.Height + 1
    Next
End Function

Function StartStarFeild(StarsNum As Long)
    ReDim Stars(StarsNum)
    StarsLow = UBound(Stars) / 4
    StarsMid = UBound(Stars) / 2
    StarsHi = StarsLow + StarsMid
    ReDim RainingText(1)
    RainingText(0).Text = Mid(FormSettings.txtRainingtext, 1, 1)
    RainingText(0).Y = 0 - FormSettings.txtTextPreview.FontSize * 8
    RainingText(0).X = Rnd * Screen.Width - 1
    With FormFullScreen
        .FontBold = FormSettings.txtTextPreview.FontBold
        .FontItalic = FormSettings.txtTextPreview.FontItalic
        .FontName = FormSettings.txtTextPreview.FontName
        .FontSize = FormSettings.txtTextPreview.FontSize
        .FontUnderline = FormSettings.txtTextPreview.FontUnderline
        .FontStrikethru = FormSettings.txtTextPreview.FontStrikethru
    End With
    Randomize
    Call StartInit
    With FormFullScreen
        If FormSettings.chkShiningStars.Value = 1 Then
            .Timer1.Interval = FormSettings.Sliders(0).Value
            .Timer1 = True
        End If
        If FormSettings.chkStars.Value = 1 Then
            .Timer2 = True
            .Timer3 = True
            .Timer4 = True
            .Timer5 = True
        End If
        If FormSettings.chkRainingText.Value = 1 Then
            .TimerRainingText = True
        End If
    End With
    FormSettings.txtRainingtext.ForeColor = FormSettings.txtTextPreview.ForeColor
End Function
