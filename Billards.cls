VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PoolBalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------
' Constants:
'-----------------------------------
Private Const m_numBalls = 16
Enum BALL_COLOURS
    BC_NONE = 0
    BC_RED = -1
    BC_YELLOW = 1
    BC_BLACK = 2
    BC_WHITE = 3
End Enum

'------------------------------
' Private data types:
'------------------------------
Private Type BALL_VARIABLES
    Texture                 As Direct3DTexture8
    Colour                  As BALL_COLOURS
    mGenerator              As D3DMATRIX
    vVelocity               As D3DVECTOR
    vSpin                   As D3DVECTOR
    FellInPocketNumber      As Integer
    numFramesSpentInPocket  As Integer
End Type

Private Type BALL_VERTEX
    XYZ As D3DVECTOR
    Normal As D3DVECTOR
    Tex1 As D3DVECTOR2
End Type
Private Const D3DFVF_BALLVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private Type SHADOW_VERTEX
    XYZ As D3DVECTOR
    Color As Long
End Type
Private Const D3DFVF_SHADOWVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'--------------------------------------------
' Variables:
'--------------------------------------------
Private m_D3DX8                     As D3DX8
Private m_D3DDevice8                As Direct3DDevice8
Private m_BallMesh                  As D3DXMesh
Private m_matWhite                  As D3DMATERIAL8

' Balls:
Private m_Balls(m_numBalls - 1)     As BALL_VARIABLES

' Shadows:
Private m_ShadowVB                  As Direct3DVertexBuffer8
Private m_numShadowVB               As Long
Private m_sizeShadowVertex          As Long
Private m_plnShadow                 As D3DPLANE

' Light direction
Private m_vLightDir                 As D3DVECTOR4

Private m_maxFramesSpentInPocket    As Integer
Private m_AnyBallInMotion           As Boolean

' Physical data common to all balls:
Private m_Mass                      As Single
Private m_Inertia                   As Single
Private m_Radius                    As Single
Private m_Elasticity                As Single

' Helpers:
Private mMtrx1 As D3DMATRIX, mMtrx2 As D3DMATRIX
Private vVctr1 As D3DVECTOR, vVctr2 As D3DVECTOR
Private a As Single, b As Single, c As Single
Private i As Long, j As Long, k As Long

'-----------------------------------------------------------------------------
' Name: NextFrame()
' Desc: Computes new generator matrices, velocity and spin vectors:
'-----------------------------------------------------------------------------
Friend Sub NextFrame()
    Dim Velocity        As Single
    Dim vTranslation    As D3DVECTOR
    Dim mTranslation    As D3DMATRIX
    Dim vRotation       As D3DVECTOR
    Dim mRotation       As D3DMATRIX
    Dim Angle           As Single
    Dim det             As Single
           
    '--------------------------
    ' Update the generators:
    '--------------------------
    m_AnyBallInMotion = False
    For i = 0 To m_numBalls - 1
        ' First, check if the ball is still "in the game". If not then go to the next ball.
        If m_Balls(i).FellInPocketNumber = 0 And m_Balls(i).numFramesSpentInPocket = 0 Then
            With m_Balls(i)
                    
                ' Generate spin caused by the ball's linear velocity:
                D3DXVec3Scale vVctr1, .vVelocity, -1
                D3DXVec3Cross vVctr2, vec3(0, -1, 0), vVctr1
                D3DXVec3Scale .vSpin, vVctr2, 1 / m_Radius
                
                ' Integrate the velocity vector into translation vector:
                D3DXVec3Scale vTranslation, .vVelocity, g_dt
                ' Translate the ball:
                With .mGenerator
                    .m41 = .m41 + vTranslation.x
                    .m42 = .m42 + vTranslation.y
                    .m43 = .m43 + vTranslation.z
                End With
                ' Reduce velocity due to friction:
                D3DXVec3Scale .vVelocity, .vVelocity, 1 - g_dt * 0.5    ' There is no special reason why there is 0.5 here - it just works well...
                
                ' If the velocity is smaller than 0.01 m/s then round
                ' it to 0.0 m/s. Truncate the spin as well.
                If D3DXVec3LengthSq(.vVelocity) < 0.0001 Then
                    .vVelocity = vec3(0, 0, 0)
                    .vSpin = vec3(0, 0, 0)
                Else
                    m_AnyBallInMotion = True
                End If
                                        
                ' Transform the spin vector into the coordinate system of the i'th ball:
                ' First create the transformation matrix:
                D3DXMatrixInverse mMtrx1, det, .mGenerator
                ' Now, integrate the spin vector into rotation vector:
                D3DXVec3Scale vRotation, .vSpin, g_dt
                ' Conduct the transformation:
                D3DXVec3TransformNormal vRotation, vRotation, mMtrx1
                ' Normalize the rotation vector:
                NormalizeD3DVECTOR vRotation, Angle, vRotation
                ' And, if the angle is greater than zero...
                If Angle > 0 Then
                    '... create a transformation matrix for this rotation:
                    D3DXMatrixRotationAxis mRotation, vRotation, Angle
                    ' Conduct the rotation transformation (rotation comes first as it is local)
                    D3DXMatrixMultiply .mGenerator, mRotation, .mGenerator
                End If
            
            End With
        Else
            With m_Balls(i)
                If .numFramesSpentInPocket <= m_maxFramesSpentInPocket Then .numFramesSpentInPocket = .numFramesSpentInPocket + 1
            End With
        End If
    Next i
End Sub


'-----------------------------------------------------------------------------
' Name: Render()
' Desc: Draws the scene
'-----------------------------------------------------------------------------
Friend Sub Render()
    Dim matBlack As D3DMATERIAL8
    Dim mTranslation As D3DMATRIX
        
    If m_D3DDevice8 Is Nothing Then Exit Sub
    
    With m_D3DDevice8
        .SetRenderState D3DRS_WRAP0, D3DWRAP_U
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetVertexShader D3DFVF_BALLVERTEX
        .SetMaterial m_matWhite
        For i = 0 To m_numBalls - 1
            If m_Balls(i).numFramesSpentInPocket <= m_maxFramesSpentInPocket Then
                ' Set the world transformation matrix:
                .SetTransform D3DTS_WORLD, m_Balls(i).mGenerator
                ' and the texture for this subset:
                .SetTexture 0, m_Balls(i).Texture
                ' Draw the mesh subset
                m_BallMesh.DrawSubset 0
            End If
        Next i

        ' Shadows:
        .SetVertexShader D3DFVF_SHADOWVERTEX
        .SetRenderState D3DRS_SPECULARENABLE, 0
        .SetRenderState D3DRS_LIGHTING, 0
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        .SetStreamSource 0, m_ShadowVB, m_sizeShadowVertex
        .SetTexture 0, Nothing
        .SetMaterial matBlack
        For i = 0 To m_numBalls - 1
            If m_Balls(i).numFramesSpentInPocket <= m_maxFramesSpentInPocket Then
                ' Set the world transformation matrix for the ball's shadow:
                With m_Balls(i).mGenerator
                    D3DXMatrixTranslation mTranslation, .m41, .m42, .m43
                End With
                D3DXMatrixShadow mMtrx1, m_vLightDir, m_plnShadow
                D3DXMatrixMultiply mMtrx2, mTranslation, mMtrx1
                .SetTransform D3DTS_WORLD, mMtrx2
                ' Draw the shadow:
                .DrawPrimitive D3DPT_TRIANGLEFAN, 0, m_numShadowVB - 2
            End If
        Next i
        .SetRenderState D3DRS_LIGHTING, 1
        .SetRenderState D3DRS_SPECULARENABLE, 1
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetRenderState D3DRS_WRAP0, 0
    End With
               
End Sub

'----------------------------------------------------
' Name: Create()
' Desc: Creates sphere mesh, sets materials etc.
'----------------------------------------------------
Friend Sub Create(ByVal inD3DDevice8 As Direct3DDevice8, ByVal inD3DX8 As D3DX8)
    Dim Angle               As Single
    Dim AdjBuffer           As D3DXBuffer
    Dim SignZ               As Single
    Dim Mesh                As D3DXMesh
    Dim VBuffer             As Direct3DVertexBuffer8
    Dim VArray()            As BALL_VERTEX
    Dim numVertices         As Long
    Dim TexPath             As String
    Dim ShadowVertices()    As SHADOW_VERTEX
    
    On Local Error GoTo RaiseError
    
    Set m_D3DDevice8 = inD3DDevice8
    Set m_D3DX8 = inD3DX8
                        
    '-----------------------------------------------
    ' Ball's radius, mass and inertia momentum:
    '-----------------------------------------------
    m_Radius = 0.05
    m_Mass = 0.1
    m_Inertia = 2 / 5 * m_Mass * m_Radius * m_Radius
    m_Elasticity = 1
        
    '-----------------------
    ' Create ball mesh:
    '-----------------------
    Set Mesh = m_D3DX8.CreateSphere(m_D3DDevice8, m_Radius, 36, 12, AdjBuffer)
    Set m_BallMesh = Mesh.CloneMeshFVF(ByVal 0, D3DFVF_BALLVERTEX, m_D3DDevice8)
    numVertices = m_BallMesh.GetNumVertices
    ReDim VArray(numVertices - 1)
    Set VBuffer = m_BallMesh.GetVertexBuffer
    D3DVertexBuffer8GetData VBuffer, 0, numVertices * Len(VArray(0)), 0, VArray(0)
        For i = 0 To numVertices - 1
            SignZ = Sgn(VArray(i).XYZ.z)
            VArray(i).Tex1.x = (-SignZ * VArray(i).XYZ.x + m_Radius) / (2 * m_Radius)
            VArray(i).Tex1.y = (VArray(i).XYZ.y + m_Radius) / (2 * m_Radius) * -1
        Next i
    D3DVertexBuffer8SetData VBuffer, 0, numVertices * Len(VArray(0)), 0, VArray(0)
    
    'Cleanup:
    Set AdjBuffer = Nothing
    Set Mesh = Nothing
    Set VBuffer = Nothing
    Erase VArray
    
    '--------------------
    ' Textures
    '--------------------
    For i = 1 To m_numBalls - 1
        TexPath = "\Ball Numbers\Ball" & Format(i) & ".bmp"
        Set m_Balls(i).Texture = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, g_AppPath + TexPath)
    Next i
    
    '-------------------------------------------------
    ' All balls share the same shiny, white material:
    '-------------------------------------------------
    With m_matWhite
        .Ambient = MakeD3DCOLORVALUE(1, 1, 1, 1)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(1, 1, 1, 1)
    End With
    
    '-----------------------
    ' Ball colours:
    '-----------------------
    ' The cue-ball is white:
    m_Balls(0).Colour = BC_WHITE
    ' The eight-ball is black:
    m_Balls(8).Colour = BC_BLACK
    ' Other balls are yellow and red:
    
    m_Balls(1).Colour = BC_RED
    m_Balls(3).Colour = BC_RED
    m_Balls(7).Colour = BC_RED
    m_Balls(9).Colour = BC_RED
    m_Balls(10).Colour = BC_RED
    m_Balls(14).Colour = BC_RED
    m_Balls(15).Colour = BC_RED
    
    m_Balls(2).Colour = BC_YELLOW
    m_Balls(4).Colour = BC_YELLOW
    m_Balls(5).Colour = BC_YELLOW
    m_Balls(6).Colour = BC_YELLOW
    m_Balls(11).Colour = BC_YELLOW
    m_Balls(12).Colour = BC_YELLOW
    m_Balls(13).Colour = BC_YELLOW
    
    
    '---------------------------
    ' Shadows' vertex buffer:
    '---------------------------
    ' NOTE: The shadow is simply a (flat) circle of the same radius as the ball mesh.
    ' ANOTHER NOTE: When renderring the contents of this buffer we will use the D3DPT_TRIANGLEFAN primitive type,
    '               thus, we don't need an index buffer.
    
    m_numShadowVB = 37      '36 is the number of "slices" in the ball mesh.
                            'The additional vertex is the middle one.
    ReDim ShadowVertices(m_numShadowVB - 1)
    
    Angle = 2 * g_PI / (m_numShadowVB - 2)
    ShadowVertices(0).XYZ = vec3(0, 0, 0)
    ShadowVertices(0).Color = D3DColorMake(0, 0, 0, 0.6)
    For i = 0 To m_numShadowVB - 2
        ShadowVertices(i + 1).XYZ = vec3(m_Radius * Sin(i * Angle), 0, m_Radius * Cos(i * Angle))
        ShadowVertices(i + 1).Color = D3DColorMake(0, 0, 0, 0.3)
    Next i
    m_sizeShadowVertex = Len(ShadowVertices(0))
    ' Create the vertex buffer:
    Set m_ShadowVB = m_D3DDevice8.CreateVertexBuffer(m_numShadowVB * m_sizeShadowVertex, 0, D3DFVF_SHADOWVERTEX, D3DPOOL_DEFAULT)
    ' Fill it:
    D3DVertexBuffer8SetData m_ShadowVB, 0, m_numShadowVB * m_sizeShadowVertex, 0, ShadowVertices(0)
    
    'Shadow plane:
    D3DXPlaneFromPoints m_plnShadow, vec3(0, 0.005, 0), vec3(1, 0.005, 0), vec3(0, 0.005, 1)
    
    'Light direction:
    m_vLightDir = GetLightDir
    
    Erase ShadowVertices
       
    Exit Sub
    
RaiseError:
    Err.Raise vbObjectError + 2, , "Unable to create Balls"
End Sub
    
'------------------------------------------
' Name: InitialPositions()
' Desc: Places the balls on the table
'       in a triangle.
'------------------------------------------
Friend Sub InitialPositions()
    Dim Inaccuracy As Single
    
    'The cue-ball:
    D3DXMatrixTranslation m_Balls(0).mGenerator, 0, m_Radius, -1
        
    ' Slightly randomize the initial position:
    Randomize
    Inaccuracy = Rnd / 4
            
    'Other balls are grouped around the eight ball in a triangular shape:
    For i = 1 To 5
        D3DXMatrixTranslation m_Balls(i).mGenerator, 0.3 - i * 0.1, m_Radius, 0.9 + Inaccuracy + 0.2
    Next i
    For i = 1 To 4
        D3DXMatrixTranslation m_Balls(i + 5).mGenerator, 0.25 - i * 0.1, m_Radius, 0.9 + Inaccuracy + 0.1
    Next i
    For i = 1 To 3
        D3DXMatrixTranslation m_Balls(i + 9).mGenerator, 0.2 - i * 0.1, m_Radius, 0.9 + Inaccuracy
    Next i
    For i = 1 To 2
        D3DXMatrixTranslation m_Balls(i + 12).mGenerator, 0.15 - i * 0.1, m_Radius, 0.9 + Inaccuracy - 0.1
    Next i
    D3DXMatrixTranslation m_Balls(m_numBalls - 1).mGenerator, 0, m_Radius, 0.9 + Inaccuracy - 0.2
    
    ' Switch the 8'th ball with the 11'th ball, so that the former is in the centre of the triangle:
    mMtrx1 = m_Balls(11).mGenerator
    m_Balls(11).mGenerator = m_Balls(8).mGenerator
    m_Balls(8).mGenerator = mMtrx1
        
    'None of the balls is initially in any pocket and all are visible, thus:
    For i = 0 To m_numBalls - 1
        With m_Balls(i)
            .FellInPocketNumber = 0
            .numFramesSpentInPocket = 0
            .vVelocity = vec3(0, 0, 0)
            .vSpin = vec3(0, 0, 0)
        End With
    Next i
    m_maxFramesSpentInPocket = 20
                     
End Sub

'------------------------------------------------------------------
' Name: FireCueBall()
' Decs: Fires the cue-ball with a given initial velocity vector
'------------------------------------------------------------------
Friend Sub FireCueBall(VelVctr As D3DVECTOR)
    m_Balls(0).vVelocity = VelVctr
    m_AnyBallInMotion = True
End Sub

'------------------------------------------------------------------
' Name: ReappearCueBall()
' Decs: Sets the number of frames spent in a pocket by the
'       cue-ball to 0.
'------------------------------------------------------------------
Friend Sub ReappearCueBall()
    m_Balls(0).numFramesSpentInPocket = 0
End Sub


'---------------------------------------------------------------------
' Name: ClearPocketNumbers()
' Decs: Sets all values in the pocket number array to "0"
'       and adds another frame to the m_numFramesSpentInPocket array.
'---------------------------------------------------------------------
Friend Sub ClearPocketNumbers()
    For i = 0 To m_numBalls - 1
        With m_Balls(i)
            If .FellInPocketNumber > 0 Then
                .numFramesSpentInPocket = m_maxFramesSpentInPocket + 1
                .FellInPocketNumber = 0
            End If
        End With
    Next i
End Sub

'-------------------------
' Read/Write Properties:
'-------------------------
Friend Property Get Generator(ByVal indx As Long) As D3DMATRIX
    If indx < m_numBalls And indx > -1 Then Generator = m_Balls(indx).mGenerator
End Property

Friend Property Let Generator(ByVal indx As Long, inGnrtrMtrx As D3DMATRIX)
    If indx < m_numBalls And indx > -1 Then m_Balls(indx).mGenerator = inGnrtrMtrx
End Property

Friend Property Get BallPosition(ByVal indx As Long) As D3DVECTOR
    If indx < m_numBalls And indx > -1 Then
        With m_Balls(indx).mGenerator
            BallPosition.x = .m41
            BallPosition.y = .m42
            BallPosition.z = .m43
        End With
    End If
End Property

Friend Property Let BallPosition(ByVal indx As Long, inPosVctr As D3DVECTOR)
    If indx < m_numBalls And indx > -1 Then
        With m_Balls(indx).mGenerator
            .m41 = inPosVctr.x
            .m42 = inPosVctr.y
            .m43 = inPosVctr.z
        End With
    End If
End Property

Friend Property Get BallVelocity(ByVal indx As Long) As D3DVECTOR
    If indx < m_numBalls And indx > -1 Then BallVelocity = m_Balls(indx).vVelocity
End Property

Friend Property Let BallVelocity(ByVal indx As Long, inVelocityVctr As D3DVECTOR)
    If indx < m_numBalls And indx > -1 Then m_Balls(indx).vVelocity = inVelocityVctr
End Property

Friend Property Get BallSpin(ByVal indx As Long) As D3DVECTOR
    If indx < m_numBalls And indx > -1 Then BallSpin = m_Balls(indx).vSpin
End Property

Friend Property Let BallSpin(ByVal indx As Long, inSpinVctr As D3DVECTOR)
    If indx < m_numBalls And indx > -1 Then m_Balls(indx).vSpin = inSpinVctr
End Property

Friend Property Get FellInPocketNumber(ByVal indx As Long) As Integer
    If indx < m_numBalls And indx > -1 Then FellInPocketNumber = m_Balls(indx).FellInPocketNumber
End Property

Friend Property Let FellInPocketNumber(ByVal indx As Long, iPocket As Integer)
    If indx < m_numBalls And indx > -1 Then m_Balls(indx).FellInPocketNumber = iPocket
End Property

'-------------------------
' Read-only Properties:
'-------------------------
Friend Property Get NumBalls() As Long
    NumBalls = m_numBalls
End Property

Friend Property Get InTheGame(ByVal indx As Long) As Boolean
    If m_Balls(indx).numFramesSpentInPocket = 0 Then InTheGame = True Else InTheGame = False
End Property

Friend Property Get AnyBallInMotion() As Boolean
    AnyBallInMotion = m_AnyBallInMotion
End Property

Friend Sub GetPhysicalConstants(ByRef retMass As Single, ByRef retInertia As Single, ByRef retRadius As Single, ByRef retElasticity As Single)
    retMass = m_Mass
    retInertia = m_Inertia
    retRadius = m_Radius
    retElasticity = m_Elasticity
End Sub

Friend Property Get Colour(ByVal indx As Long) As BALL_COLOURS
    Colour = m_Balls(indx).Colour
End Property

'-----------------------------------------------------------------------------
' Name: Class_Terminate()
' Desc: Called when an instance of this class is destroyed
'-----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Erase m_Balls
    
    Set m_ShadowVB = Nothing
    Set m_BallMesh = Nothing
    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing
End Sub




'===============================
'     HELPER FUNCTIONS:
'===============================

Private Sub NormalizeD3DVECTOR(retV As D3DVECTOR, retModule As Single, vVctr As D3DVECTOR)
    Dim module As Single
    module = Sqr(vVctr.x * vVctr.x + vVctr.y * vVctr.y + vVctr.z * vVctr.z)
    If module > 0 Then retV.x = vVctr.x / module: retV.y = vVctr.y / module: retV.z = vVctr.z / module
    retModule = module
End Sub

