VERSION 5.00
Begin VB.Form FormFullScreen 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerRainingText 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3240
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2640
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2040
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1440
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   840
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   240
      Top             =   480
   End
End
Attribute VB_Name = "FormFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1 = False
    Timer2 = False
    Timer3 = False
    Timer4 = False
    Timer5 = False
    TimerRainingText = False
    FormSettings.txtRainingtext.ForeColor = &H80000007
    RainIndex = 1
    RainingTextLoopIndex = 0
    FormFullScreen.Hide
    StopLoop = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormFullScreen.Hide
    Timer1 = False
    Timer2 = False
    Timer3 = False
    Timer4 = False
    Timer5 = False
    TimerRainingText = False
    FormSettings.txtRainingtext.ForeColor = &H80000007
    RainIndex = 1
    RainingTextLoopIndex = 0
    FormFullScreen.Hide
    StopLoop = False
End Sub

Private Sub Timer1_Timer()
    FormFullScreen.Circle (LastStar(0), LastStar(1)), FormSettings.Sliders(5).Value, &H80000007
    LastStar(0) = Rnd * Screen.Width + 1
    LastStar(1) = Rnd * Screen.Height + 1
    FormFullScreen.Circle (LastStar(0), LastStar(1)), FormSettings.Sliders(5).Value, FormSettings.cmdShiningStarsColor.BackColor
End Sub

Private Sub Timer2_Timer()
    Dim i As Long
    For i = 0 To StarsLow - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), &H80000007
    Next
    For i = 0 To StarsLow - 1
        Stars(i).Y = Stars(i).Y + FormSettings.lblSliders(1)
        If Stars(i).Y > Screen.Height Then
            Stars(i).X = Rnd * Screen.Width + 1
            Stars(i).Y = 0
        End If
    Next
    For i = 0 To StarsLow - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), FormSettings.cmdSliders(0).BackColor
    Next
End Sub

Private Sub Timer3_Timer()
    Dim i As Long
    For i = StarsLow To StarsMid - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), &H80000007
    Next
    For i = StarsLow To StarsMid - 1
        Stars(i).Y = Stars(i).Y + FormSettings.lblSliders(2)
        If Stars(i).Y > Screen.Height Then
            Stars(i).X = Rnd * Screen.Width + 1
            Stars(i).Y = 0
        End If
    Next
    For i = StarsLow To StarsMid - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), FormSettings.cmdSliders(1).BackColor
    Next
End Sub

Private Sub Timer4_Timer()
    Dim i As Long
    For i = StarsMid To StarsHi - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), &H80000007
    Next
    For i = StarsMid To StarsHi - 1
        Stars(i).Y = Stars(i).Y + FormSettings.lblSliders(3)
        If Stars(i).Y > Screen.Height Then
            Stars(i).X = Rnd * Screen.Width + 1
            Stars(i).Y = 0
        End If
    Next
    For i = StarsMid To StarsHi - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), FormSettings.cmdSliders(2).BackColor
    Next
End Sub

Private Sub Timer5_Timer()
    Dim i As Long
    For i = StarsHi To UBound(Stars) - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), &H80000007
    Next
    For i = StarsHi To UBound(Stars) - 1
        Stars(i).Y = Stars(i).Y + FormSettings.lblSliders(4)
        If Stars(i).Y > Screen.Height Then
            Stars(i).X = Rnd * Screen.Width + 1
            Stars(i).Y = 0
        End If
    Next
    For i = StarsHi To UBound(Stars) - 1
        FormFullScreen.PSet (Stars(i).X, Stars(i).Y), FormSettings.cmdSliders(3).BackColor
    Next
End Sub

Private Sub TimerRainingText_Timer()
    Dim i As Long
    For i = 0 To UBound(RainingText) - 1 ' clear
        Me.ForeColor = &H80000007
        Me.CurrentX = RainingText(i).X
        Me.CurrentY = RainingText(i).Y
        Me.Print RainingText(i).Text
    Next
    For i = 0 To UBound(RainingText) - 1 ' draw
        RainingText(i).Y = RainingText(i).Y + FormSettings.sliderRainingTextSpeed.Value
        Me.ForeColor = FormSettings.txtRainingtext.ForeColor
        Me.CurrentX = RainingText(i).X
        Me.CurrentY = RainingText(i).Y
        Me.Print RainingText(i).Text
    Next
    Steps = Steps + 1
    If Steps < FormSettings.sliderRainingTextStep.Value = 0 And UBound(RainingText) < Len(FormSettings.txtRainingtext) + 1 Then
        RainIndex = RainIndex + 1
        If RainIndex = Len(FormSettings.txtRainingtext) Then
            LastLetter = True
        End If
        ReDim Preserve RainingText(RainIndex)
        RainingText(RainIndex - 1).Text = Mid(FormSettings.txtRainingtext, RainIndex, 1)
        RainingText(RainIndex - 1).X = Rnd * Me.Width - 1
        RainingText(RainIndex - 1).Y = 0 - FormSettings.txtTextPreview.FontSize * 8
        Steps = 0
    End If
    If LastLetter And RainingText(RainIndex - 1).Y >= Me.Height Then
        If FormSettings.chkRaingTextLoop.Value = 1 Then
            ReDim RainingText(1)
            RainingText(0).Text = Mid(FormSettings.txtRainingtext, 1, 1)
            RainingText(0).Y = 0 - FormSettings.txtTextPreview.FontSize * 8
            RainingText(0).X = Rnd * Screen.Width - 1
            RainIndex = 1
            LastLetter = False
        Else
            TimerRainingText = False
        End If
    End If
End Sub
