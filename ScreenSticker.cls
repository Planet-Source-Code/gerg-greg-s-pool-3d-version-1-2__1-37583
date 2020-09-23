VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScreenSticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' References to DirectX objects:
Private m_D3DDevice8        As Direct3DDevice8
Private m_D3DX8             As D3DX8
Private m_Texture           As Direct3DTexture8

' Sticker characteristics:
Private m_Rect              As RECT
Private m_vRotCentre        As D3DVECTOR2
Private m_Rotation          As Single
Private m_vScale            As D3DVECTOR2
Private m_vTranslate        As D3DVECTOR2
Private m_Color             As Long
Private m_CurrentColor      As Long
Private m_bVisible          As Boolean

' Flashing colors:
Private m_bFlashing         As Boolean
Private m_FlashClr1         As D3DCOLORVALUE
Private m_FlashClr2         As D3DCOLORVALUE
Private m_FlashFreq         As Single

' Mouse cursor:
Private m_bUnderCursor      As Boolean

'-------------------------------------
' Name: Setup()
' Desc:
'-------------------------------------
Friend Sub Setup(ByRef TexSrc As String, _
                 ByVal PixelFormat As CONST_D3DFORMAT, _
                 ByVal SpriteLeft As Single, _
                 ByVal SpriteTop As Single, _
                 ByVal SpriteWidth As Single, _
                 ByVal SpriteHeight As Single, _
                 ByVal SpriteColor As Long, _
                 ByVal inD3DDevice8 As Direct3DDevice8, _
                 ByVal inD3DX8 As D3DX8)
    
   
    ' Copy the references to DirectX objects:
    Set m_D3DDevice8 = inD3DDevice8
    Set m_D3DX8 = inD3DX8
    
    ' Create the sprite bounding rectangle:
    With m_Rect
        .Right = SpriteWidth
        .bottom = SpriteHeight
    End With
    
    ' And the translation vector:
    With m_vTranslate
        .y = SpriteTop
        .x = SpriteLeft
    End With
    
    ' Copy the sprite's color:
    m_Color = SpriteColor
    
    ' The scaling vector is by default (1,1):
    m_vScale.x = 1: m_vScale.y = 1
    
    ' The centre of rotation is by default
    ' in the middle of m_Rect:
    m_vRotCentre.x = m_Rect.Left + (m_Rect.Right - m_Rect.Left) / 2
    m_vRotCentre.y = m_Rect.Top + (m_Rect.bottom - m_Rect.Top) / 2
    
    ' Rotation is initially set to zero.
    m_Rotation = 0
        
    ' Create the texture
    Set m_Texture = m_D3DX8.CreateTextureFromFileEx(m_D3DDevice8, TexSrc, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, PixelFormat, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &H0, ByVal 0, ByVal 0)
    
End Sub

'---------------------------------------------------
' Name: Draw()
' Desc: Draws the screen sticker image using
'       a D3DXSprite object provided by the caller.
'---------------------------------------------------
Friend Sub Draw(ByVal inSprite As D3DXSprite)
    If m_bUnderCursor Then FlashColors Else m_CurrentColor = m_Color
    inSprite.Draw m_Texture, m_Rect, m_vScale, m_vRotCentre, m_Rotation, m_vTranslate, m_CurrentColor
End Sub

Friend Function UnderCursor(ByVal CurX As Single, ByVal CurY As Single) As Boolean
    If (CurX >= m_Rect.Left + m_vTranslate.x And CurX <= m_Rect.Right + m_vTranslate.x) _
        And (CurY >= m_Rect.Top + m_vTranslate.y And CurY <= m_Rect.bottom + m_vTranslate.y) Then
        UnderCursor = True
        m_bUnderCursor = True
    Else
        UnderCursor = False
        m_bUnderCursor = False
    End If
End Function

Friend Sub SetupColorFlashing(ByRef Clr1 As D3DCOLORVALUE, ByRef Clr2 As D3DCOLORVALUE, ByVal Freq As Single)
    m_FlashClr1 = Clr1
    m_FlashClr2 = Clr2
    m_FlashFreq = Freq
    m_bFlashing = True
End Sub

Friend Sub StopFlashingColors()
    m_bFlashing = False
End Sub

'---------------------------------------------------
' Property: Visible()
' Type: Read/write
' Desc: Returns true if the sticker is visible,
'       false if it's not
'---------------------------------------------------
Friend Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Friend Property Let Visible(bVal As Boolean)
    m_bVisible = bVal
End Property

Friend Property Get StillFlashing() As Boolean
    StillFlashing = m_bFlashing
End Property

'--------------------------------------------------------------
' Name: FlashColors()
' Desc: Computes the m_CurrentColor value by linear
'       interpolation of m_FlashClr1 and m_FlashClr2.
'--------------------------------------------------------------
Private Sub FlashColors()
    Dim Blend As Single
    Dim ClrFinal As D3DCOLORVALUE
    
    Blend = (Sin(m_FlashFreq * Timer) + 1) / 2
    D3DXColorLerp ClrFinal, m_FlashClr1, m_FlashClr2, Blend
    m_CurrentColor = D3DColorMake(ClrFinal.r, ClrFinal.g, ClrFinal.b, ClrFinal.a)
End Sub

'----------------------------------
' Name: Class_Terminate()
' Desc: Clears all objects.
'----------------------------------
Private Sub Class_Terminate()
    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing
    Set m_Texture = Nothing
End Sub
