VERSION 5.00
Begin VB.Form frmPool 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Greg's Pool 3D - Prototype Version 1.2"
   ClientHeight    =   8025
   ClientLeft      =   945
   ClientTop       =   2625
   ClientWidth     =   11475
   Icon            =   "frmPool.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   Begin VB.PictureBox picPower 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   240
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Move the diamond shaped indicator to set the initial velocity of cue-ball."
      Top             =   6600
      Width           =   3750
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      ForeColor       =   &H0080FF80&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   763
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11475
   End
   Begin VB.Label lblShoot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    Shoot    "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4755
      TabIndex        =   3
      ToolTipText     =   "Click here to launch the cue-ball"
      Top             =   6840
      Width           =   1905
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   510
      Left            =   7935
      TabIndex        =   2
      ToolTipText     =   "Current player"
      Top             =   6840
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPanel 
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Left            =   120
      Picture         =   "frmPool.frx":08CA
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   11400
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New              F2"
      End
      Begin VB.Menu mnuGameBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Exit               Esc"
      End
   End
   Begin VB.Menu mnuStgs 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSV 
         Caption         =   "Sound &Volume"
         Begin VB.Menu mnuSVMute 
            Caption         =   "&Mute"
         End
         Begin VB.Menu mnuSVQuiet 
            Caption         =   "&Quiet"
         End
         Begin VB.Menu mnuSVAve 
            Caption         =   "&Average"
         End
         Begin VB.Menu mnuSVLoud 
            Caption         =   "&Loud"
         End
      End
      Begin VB.Menu mnuTG 
         Caption         =   "Toggle Cameras                                  Tab"
      End
      Begin VB.Menu mnuN8BP 
         Caption         =   "Name 8'th ball pocket"
      End
      Begin VB.Menu mnuTCL 
         Caption         =   "Toggle cameras after launching the cue-ball"
      End
   End
End
Attribute VB_Name = "frmPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POWER_INDICATOR
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
    MaxLeft     As Single
    MinLeft     As Single
End Type

Private m_CurrentX          As Single
Private m_CurrentY          As Single

' Cursors:
Private m_LButtonCur        As Integer
Private m_RButtonCur        As Integer
Private m_NoButtonCur       As Integer

Private m_PwrIndctr         As POWER_INDICATOR
Private m_Power             As Single
Private m_MaxPower          As Single

Private m_PicPowerRange     As StdPicture
Private m_PicPwrIndctr      As StdPicture
Private m_PicPwrIndctrMask  As StdPicture

Private Sub Form_Load()
    m_MaxPower = 7
    
    Set m_PicPwrIndctr = LoadPicture(g_AppPath & "\PowerIndicator.bmp")
    Set m_PicPwrIndctrMask = LoadPicture(g_AppPath & "\PowerIndicatorMask.bmp")
    Set m_PicPowerRange = LoadPicture(g_AppPath & "\PowerRange.bmp")
    
    With Me
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        .ScaleHeight = 750
        .ScaleWidth = 1000
    End With
    
    ' Main picture box:
    Picture1.Height = 600
    
    ' Controll panel:
    With imgPanel
        .Left = 0
        .Top = 600
        .Width = 1000
        .Height = 150
    End With
    
    ' Power setting picture box:
    With picPower
        .ScaleHeight = 80
        .ScaleWidth = 250
        .Left = 20
        .Top = 620
        .Height = 114
        .Width = 330
        .AutoRedraw = True
        .PaintPicture m_PicPowerRange, 0, 0, .ScaleWidth, .ScaleHeight
        .AutoRedraw = False
    End With
    
    ' Power indicator
    With m_PwrIndctr
        .Height = 34
        .Width = 20
        .MaxLeft = 219
        .MinLeft = 9
        .Left = (.MaxLeft - .MinLeft) / 2 + .MinLeft
        .Top = 14
    End With
    
    ' Initial power value:
    m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
           
    ' The "shoot button" label:
    With lblShoot
        .FontSize = 24
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = 640
    End With
       
    ' The label displaying current player number:
    With lblPlayer
        .FontSize = 24
        .Left = Me.ScaleWidth - .Width - 20
        .Top = 620
    End With
      
    'Sounds:
    mnuSVAve.Checked = True
    mnuSVAve_Click
    
    'Default settings:
    mnuTCL_Click
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace: FireCueBall m_Power
        Case vbKeyEscape: Form_Unload 0
        Case vbKeyTab: ToggleCameras
        Case vbKeyF2: StartNewGame
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopGame
    Set m_PicPwrIndctrMask = Nothing
    Set m_PicPwrIndctr = Nothing
    Set m_PicPowerRange = Nothing

    End
End Sub

Private Sub lblShoot_Click()
    FireCueBall m_Power
End Sub

'--------------------
'     GAME MENU:
'--------------------
Private Sub mnuGameExit_Click()
    Form_Unload 0
End Sub

Private Sub mnuGameNew_Click()
    StartNewGame
End Sub

'---------------------
'   SETTINGS MENU:
'---------------------
Private Sub mnuTCL_Click()
    If mnuTCL.Checked Then
        mnuTCL.Checked = False
        ToggleCamAfterLaunch False
    Else
        mnuTCL.Checked = True
        ToggleCamAfterLaunch True
    End If
End Sub

Private Sub mnuN8BP_Click()
    If mnuN8BP.Checked Then
        mnuN8BP.Checked = False
        NameEightBallPocket False
    Else
        mnuN8BP.Checked = True
        NameEightBallPocket True
    End If
End Sub

Private Sub mnuTG_Click()
    ToggleCameras
End Sub

Private Sub mnuSVMute_Click()
    mnuSVMute.Checked = True
    mnuSVQuiet.Checked = False
    mnuSVAve.Checked = False
    mnuSVLoud.Checked = False
    SetSoundVolumeBase -4000
End Sub

Private Sub mnuSVQuiet_Click()
    mnuSVMute.Checked = False
    mnuSVQuiet.Checked = True
    mnuSVAve.Checked = False
    mnuSVLoud.Checked = False
    SetSoundVolumeBase -3000
End Sub

Private Sub mnuSVAve_Click()
    mnuSVMute.Checked = False
    mnuSVQuiet.Checked = False
    mnuSVAve.Checked = True
    mnuSVLoud.Checked = False
    SetSoundVolumeBase -2000
End Sub

Private Sub mnuSVLoud_Click()
    mnuSVMute.Checked = False
    mnuSVQuiet.Checked = False
    mnuSVAve.Checked = False
    mnuSVLoud.Checked = True
    SetSoundVolumeBase -1000
End Sub


Private Sub picPower_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With m_PwrIndctr
            .Left = x - .Width / 2
            If .Left < .MinLeft Then .Left = .MinLeft
            If .Left > .MaxLeft Then .Left = .MaxLeft
        End With
        With picPower
            .Cls
            .PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
            .PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
        End With
        m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
    End If
End Sub

Private Sub picPower_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With m_PwrIndctr
            .Left = x - .Width / 2
            If .Left < .MinLeft Then .Left = .MinLeft
            If .Left > .MaxLeft Then .Left = .MaxLeft
        End With
        With picPower
            .Cls
            .PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
            .PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
        End With
        m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
    End If
End Sub

Private Sub picPower_Paint()
    ' Power indicator:
    picPower.PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
    picPower.PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_CurrentX = x
    m_CurrentY = y
    MouseEventHandler Button, x, y, 0, 0
    
    ' Set the cursor
    If Button = 1 Then
        Picture1.MousePointer = m_LButtonCur
    ElseIf Button = 2 Then
        Picture1.MousePointer = m_RButtonCur
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ShiftX As Single
    Dim ShiftY As Single
            
    ShiftX = x - m_CurrentX
    ShiftY = y - m_CurrentY
    m_CurrentX = x: m_CurrentY = y
    If m_CurrentX >= 0 And m_CurrentX <= Picture1.ScaleWidth And m_CurrentY >= 0 And m_CurrentY <= Picture1.ScaleHeight Then
        ShiftX = ShiftX / Picture1.ScaleWidth * 0.2
        ShiftY = ShiftY / Picture1.ScaleHeight * 0.2
        MouseEventHandler Button, x, y, ShiftX, ShiftY
    End If
    
    ' Set the cursor
    If Button = 1 Then
        Picture1.MousePointer = m_LButtonCur
    ElseIf Button = 2 Then
        Picture1.MousePointer = m_RButtonCur
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.MousePointer = m_NoButtonCur
End Sub

Friend Sub SetCursorType(Optional NoButtonCur As Integer, Optional LButtonCur As Integer, Optional RButtonCur As Integer)
    m_NoButtonCur = NoButtonCur
    m_LButtonCur = LButtonCur
    m_RButtonCur = RButtonCur
    Picture1.MousePointer = m_NoButtonCur
End Sub
