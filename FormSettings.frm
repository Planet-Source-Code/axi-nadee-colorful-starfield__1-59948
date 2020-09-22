VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Axi's Starfield"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "FormSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmEffects 
      Caption         =   "Options"
      Height          =   2055
      Left            =   5640
      TabIndex        =   32
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox chkRaingTextLoop 
         Caption         =   "Loop Raining Text"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkShiningStars 
         Caption         =   "Shining Stars"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkStars 
         Caption         =   "Star Feild"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkRainingText 
         Caption         =   "Raining Text"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4920
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmRainingText 
      Caption         =   "Raining Text"
      Height          =   4215
      Left            =   2880
      TabIndex        =   21
      Top             =   960
      Width           =   2655
      Begin ComctlLib.Slider sliderRainingTextSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         Max             =   360
         SelStart        =   1
         TickFrequency   =   36
         Value           =   1
      End
      Begin VB.TextBox txtRainingtext 
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "FormSettings.frx":08CA
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "Color"
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtTextPreview 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1005
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "FormSettings.frx":08FD
         Top             =   3120
         Width           =   2415
      End
      Begin ComctlLib.Slider sliderRainingTextStep 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         Max             =   500
         SelStart        =   500
         TickFrequency   =   50
         Value           =   500
      End
      Begin VB.Label lblSliders 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   31
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblT 
         Caption         =   "Spaces Between Letters:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblSliders 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   28
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblT 
         Caption         =   "Speed:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame frmShiningStars 
      Caption         =   "Shining Stars Settings"
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   2655
      Begin VB.CommandButton cmdShiningStarsColor 
         BackColor       =   &H000080FF&
         Caption         =   "Color"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Speed"
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   10
         Min             =   1
         Max             =   1000
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   100
         Value           =   1
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Size"
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   327682
         LargeChange     =   10
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   10
         Value           =   1
      End
      Begin VB.Label lblSliders 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   20
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblSliders 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   18
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame frmLayers 
      Caption         =   "Stars Layers Speed"
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdSliders 
         BackColor       =   &H00FF0000&
         Caption         =   "4"
         Height          =   375
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSliders 
         BackColor       =   &H0000FF00&
         Caption         =   "3"
         Height          =   375
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSliders 
         BackColor       =   &H000000FF&
         Caption         =   "2"
         Height          =   375
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSliders 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1"
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   1935
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   360
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   36
         Value           =   1
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   1935
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   360
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   36
         Value           =   1
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   1935
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   360
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   36
         Value           =   1
      End
      Begin ComctlLib.Slider Sliders 
         Height          =   1935
         Index           =   4
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   360
         SelStart        =   1
         TickStyle       =   2
         TickFrequency   =   36
         Value           =   1
      End
      Begin VB.Label lblSliders 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSliders 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   10
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSliders 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblSliders 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   375
      End
   End
   Begin VB.Frame frmNumStars 
      Caption         =   "Number of Stars"
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtStarsNum 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1037
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "200"
         Top             =   270
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   5640
      Picture         =   "FormSettings.frx":090C
      Stretch         =   -1  'True
      ToolTipText     =   "This Software was made by Axi from USHASOFT 2005"
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFont_Click()
    cmdlg.FontBold = txtTextPreview.FontBold
    cmdlg.FontItalic = txtTextPreview.FontItalic
    cmdlg.FontName = txtTextPreview.Font
    cmdlg.FontSize = txtTextPreview.FontSize
    cmdlg.FontStrikethru = txtTextPreview.FontStrikethru
    cmdlg.FontUnderline = txtTextPreview.FontUnderline
    cmdlg.Flags = cdlCFBoth
    cmdlg.ShowFont
    txtTextPreview.FontBold = cmdlg.FontBold
    txtTextPreview.FontItalic = cmdlg.FontItalic
    txtTextPreview.FontName = cmdlg.FontName
    txtTextPreview.FontSize = cmdlg.FontSize
    txtTextPreview.FontStrikethru = cmdlg.FontStrikethru
    txtTextPreview.FontUnderline = cmdlg.FontUnderline
End Sub

Private Sub cmdFontColor_Click()
    cmdlg.Color = txtTextPreview.ForeColor
    cmdlg.ShowColor
    txtTextPreview.ForeColor = cmdlg.Color
End Sub

Private Sub cmdShiningStarsColor_Click()
    cmdlg.Color = cmdShiningStarsColor.BackColor
    cmdlg.ShowColor
    cmdShiningStarsColor.BackColor = cmdlg.Color
End Sub

Private Sub cmdSliders_Click(Index As Integer)
    cmdlg.Color = cmdSliders(Index).BackColor
    cmdlg.ShowColor
    cmdSliders(Index).BackColor = cmdlg.Color
End Sub

Private Sub cmdStart_Click()
    Call StartStarFeild(txtStarsNum)
    Steps = 0
    FormFullScreen.Show (1)
End Sub

Private Sub Form_Load()
    Sliders(0).Value = 20
    Sliders(1).Value = 40
    Sliders(2).Value = 70
    Sliders(3).Value = 85
    Sliders(4).Value = 100
    Sliders(5).Value = 10
    sliderRainingTextStep.Value = 5
    sliderRainingTextSpeed.Value = 200
    RainIndex = 1
    RainingTextLoopIndex = 0
    StopLoop = False
End Sub

Private Sub sliderRainingTextSpeed_Change()
    lblSliders(6) = sliderRainingTextSpeed.Value
End Sub

Private Sub sliderRainingTextSpeed_Click()
    lblSliders(6) = sliderRainingTextSpeed.Value
End Sub

Private Sub sliderRainingTextSpeed_Scroll()
    lblSliders(6) = sliderRainingTextSpeed.Value
End Sub

Private Sub sliderRainingTextStep_Change()
    lblSliders(7) = sliderRainingTextStep.Value
End Sub

Private Sub sliderRainingTextStep_Click()
    lblSliders(7) = sliderRainingTextStep.Value
End Sub

Private Sub sliderRainingTextStep_Scroll()
    lblSliders(7) = sliderRainingTextStep.Value
End Sub

Private Sub Sliders_Change(Index As Integer)
    lblSliders(Index).Caption = Sliders(Index).Value
End Sub

Private Sub Sliders_Click(Index As Integer)
    lblSliders(Index).Caption = Sliders(Index).Value
End Sub

Private Sub Sliders_Scroll(Index As Integer)
    lblSliders(Index).Caption = Sliders(Index).Value
End Sub
