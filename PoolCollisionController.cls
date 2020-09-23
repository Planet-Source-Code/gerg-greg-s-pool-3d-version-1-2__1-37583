VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PoolCollisionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+-----------------------------------------------------------------------+
'|   GENERAL NOTES:                                                      |
'|                                                                       |
'|   - A capital "C" after "m_" or at the beginning of a variable name   |
'|     (eg. "m_CBillBillMomentum") stands for "Collision".               |
'|   - A capital "B" after "m_" or at the beginning of a variable name   |
'|     (eg. "m_BMass") stands for "Ball".                             |
'|   - A capital "T" after "m_" or at the beginning of a variable name   |
'|     (eg. "m_TFriction") stands for "Table".                           |
'|                                                                       |
'+-----------------------------------------------------------------------+

'-------------------------------------------------
' Reference to objects taking part in collisions.
'-------------------------------------------------
Private m_PoolBalls                          As PoolBalls
Private m_Table                             As PoolTable

'-----------------
' Physical data:
'-----------------
' The balls:
Private m_vBPos()                           As D3DVECTOR    ' Array of balls' position vectors
Private m_vBVelocities()                    As D3DVECTOR    ' Array of balls' linear velocity vectors
Private m_vBSpins()                         As D3DVECTOR    ' Array of balls' spin (angular velocity) vectors
Private m_BMass                             As Single       ' Ball's mass
Private m_BInertia                          As Single       ' Ball's inertia (around any axis).
Private m_BRadius                           As Single       ' Ball radius.
Private m_2BRadSqrd                         As Single       ' Ball radius, squared and multiplied by two.
Private m_BElasticity                       As Single
' The table:
Private m_TFriction                         As Single       ' Table friction coefficient (currently not used).
Private m_TBumperElasticity                 As Single       ' Table elasticity coefficient (used for collision reactions).
Private m_TBumpers(1 To 6, 1 To 5)          As D3DVECTOR    ' Array of points (vectors) representing bumpers.
Private m_TBumperNorms(1 To 6, 1 To 5)      As D3DVECTOR    ' Array of normal vectors for the bumers.

'----------------------------
' Ball "presence" list:
'----------------------------
Private m_bBInTheGame()                     As Boolean      ' Each value of this array represents a Ball.
                                                            ' If a value is "true" the corresponding ball is still on the table.
'-----------------------
' Pockets:
'-----------------------
Private m_vTPockets(1 To 6)                  As D3DVECTOR    ' Array of pocket positions.
Private m_TPocketRadSqrd                    As Single       ' Squared radius of table pockets.

'-----------------------
' Collisions detected:
'-----------------------
Private m_CBillBillMomentum                 As Single       ' The scalar value of momentum for the collision
                                                            ' between two balls.
                                                            ' (currently used only for setting collision sound volume)
Private m_CBillTableMomentum                As Single       ' The scalar value of momentum for the collision
                                                            ' between a ball and the table.
                                                            ' (currently used only for setting collision sound volume)
Private m_bPocketHitDetected                As Boolean      ' True when a ball falls into a pocket.

'------------
' Counters:
'------------
Private m_numPoolBalls                       As Integer      ' Yeah, You guessed it: the total number of balls in the game.
Private Const m_numTableBumpers% = 6                        ' The number of table bumpers

'------------------------------
' The colour of the first ball
' touched by the cue-ball:
'------------------------------
Private m_FirstBallClr                      As BALL_COLOURS
 
'-----------
' Helpers:
'-----------
Private vVctr1 As D3DVECTOR, vVctr2 As D3DVECTOR
Private a As Single, b As Single, c As Single
Private i As Long, j As Long, k As Long

Friend Sub Setup(ByVal inBalls As PoolBalls, ByVal inTable As PoolTable)
    Set m_PoolBalls = inBalls
    Set m_Table = inTable
    
    ' Don't bother if there are no balls yet.
    If m_PoolBalls Is Nothing Then GoTo RaiseError
    
    ' Get the relevant data from the Balls object:
    With m_PoolBalls
        ' We will need the total number of balls...
        m_numPoolBalls = .NumBalls
        '... and their physical properties:
        .GetPhysicalConstants m_BMass, m_BInertia, m_BRadius, m_BElasticity
    End With
    ' Resize the relevant arrays:
    ReDim m_vBPos(m_numPoolBalls - 1)
    ReDim m_vBVelocities(m_numPoolBalls - 1)
    ReDim m_vBSpins(m_numPoolBalls - 1)
    ReDim m_bBInTheGame(m_numPoolBalls - 1)
    ' Compute the m_2BRadSqrd variable:
    m_2BRadSqrd = (2 * m_BRadius) ^ 2
    
    ' Don't bother if there is no table yet.
    If m_Table Is Nothing Then GoTo RaiseError
    
    ' Get the relevant data from the PoolTable object:
    With m_Table
        ' We will need the bumpers,
        .GetBumpers m_TBumpers
        ' their physical properties,
        .GetPhysicalConstants m_TBumperElasticity, m_TFriction
        ' and the pockets:
        .GetPockets m_vTPockets, a      'NOTE: We've used the "a" variable to get the pockets' radius
    End With                            '      because we don't really need the radius as such.
                                        '      We will use "a" to compute m_TPocketRadSqrd:
    m_TPocketRadSqrd = a * a
    
    
    ' Bumper architecture:
    ' A typical bumper looks like this:
    '
    '               1/2                             5
    '                  \                           /
    '                   \                         /
    '                    3 --------------------- 4
    '
    ' NOTE: The numbers refer to vertices (vectors) in the m_TBumpers array.
    ' ANOTHER NOTE: The first and second vertices are actually the same vertex.
    '               This way we will have five normals and five vertices. This will proove
    '               very convenient, when the bumper vertices and bumper normal vectors will be
    '               used together for collision detection.
    '
    '
    '
    ' Compute the bumpers' normal vectors.
    ' This is done by calculating the cross product of a vector running along the bumper's edge and the
    ' Y-axis:
    For i = 1 To m_numTableBumpers
        For j = 2 To 4
            D3DXVec3Subtract vVctr1, m_TBumpers(i, j + 1), m_TBumpers(i, j)     ' vVctr1 is the vector running along the bumper's edge
            vVctr1.y = 0    ' "Flatten" the vector, just to be sure...
            D3DXVec3Cross vVctr1, vVctr1, vec3(0, 1, 0)     ' This is the cross product with the y-axis.
            D3DXVec3Normalize m_TBumperNorms(i, j), vVctr1  ' The normal vector should be of unit length.
        Next j
        ' The first and the last normal vectors are obtained with a cross product
        ' of the third normal and the y-axis (pointing upwards or downwards).
        D3DXVec3Cross m_TBumperNorms(i, 1), m_TBumperNorms(i, 3), vec3(0, 1, 0)
        D3DXVec3Cross m_TBumperNorms(i, 5), m_TBumperNorms(i, 3), vec3(0, -1, 0)
    Next i
        
    m_FirstBallClr = BC_NONE
    
    Exit Sub
    
RaiseError:
    Err.Raise vbObjectError + 3, , "Unable to create the Collision Controller"
End Sub

Friend Sub RunCollisions()
    
    ' First, get the newest generators, velocity and spin vectors into private arrays:
    With m_PoolBalls
        For i = 0 To m_numPoolBalls - 1
            m_vBPos(i) = .BallPosition(i)
            m_vBVelocities(i) = .BallVelocity(i)
            m_vBSpins(i) = .BallSpin(i)
            m_bBInTheGame(i) = .InTheGame(i)
        Next i
    End With
    
    ' We haven't detected any collisions yet, thus:
    m_CBillBillMomentum = 0
    m_CBillTableMomentum = 0
    m_bPocketHitDetected = False
    
    ' Now, run the collision tests and eventual reactions:
    ' First, for follisions between balls...
    CDnR_BallBall
    ' ... then for collisions between the balls and the table:
    CDnR_BallTable   ' This sub also looks for pocket hits
    ' NOTE: CDnR stands for "Collision Detection and Responce".
           
    ' Finally, update the generators, velocity and spin vectors, from the private arrays:
    With m_PoolBalls
        For i = 0 To m_numPoolBalls - 1
            .BallPosition(i) = m_vBPos(i)
            .BallVelocity(i) = m_vBVelocities(i)
            .BallSpin(i) = m_vBSpins(i)
        Next i
    End With
    
End Sub

'-------------------------------------------------------------------------------
' Name: CDnR_BallBall()
' Desc: Detects collisions between balls and generates collision responce:
'-------------------------------------------------------------------------------
Private Sub CDnR_BallBall()
    Dim vBall_i      As D3DVECTOR    ' Position vector of the i'th ball
    Dim vBall_j      As D3DVECTOR    ' Position vector of the j'th ball
    Dim vVel_i          As D3DVECTOR    ' Velocity vector of the i'th ball
    Dim vVel_j          As D3DVECTOR    ' Velocity vector of the j'th ball
    Dim vDist           As D3DVECTOR    ' A vector linkng the centres of both balls.
    Dim Delta           As Single       ' A utility variable.
    Dim Ct1             As Single       ' First possible point in time, when a collision occurs.
    Dim Ct2             As Single       ' Second possible point in time, when a collision occurs.
    Dim Ct              As Single       ' The real collision point in time.
    Dim vCNorm          As D3DVECTOR    ' Collision Normal vector.
    Dim vCP_i           As D3DVECTOR    ' Collision Point relative to the centre of the i'th ball
    Dim vCP_j           As D3DVECTOR    ' Collision Point relative to the centre of the j'th ball
    Dim vVel_ij         As D3DVECTOR    ' Relative velocity between the i'th and j'th ball.
    Dim vCPVel_i        As D3DVECTOR    ' Velocity vector of the i'th ball AT POINT OF COLLISION.
    Dim vCPVel_j        As D3DVECTOR    ' Velocity vector of the j'th ball AT POINT OF COLLISION.
    Dim vCPVel_ij       As D3DVECTOR    ' Relative velocity between the i'th and j'th ball at collision point.
    Dim CMmntm          As Single       ' The momentum of collision.
       
    ' Run through all balls from the back of the array to its front:
    For i = m_numPoolBalls - 1 To 0 Step -1
        ' Don't bother looking for collisions, if the ball is not on the table any more:
        If Not m_bBInTheGame(i) Then GoTo NextI
        ' Copy the i'th ball's position and velocity
        ' to local variables:
        vBall_i = m_vBPos(i)
        vVel_i = m_vBVelocities(i)
        ' For every i'th ball iterate through all other balls in serch for collisions:
        For j = m_numPoolBalls - 1 To 0 Step -1
            ' Don't bother looking for collisions, if the j'th ball is not on the table any more, or if j = i:
            If Not m_bBInTheGame(j) Or j = i Then GoTo NextJ
            ' Copy the j'th ball's position and velocity
            ' to local variables:
            vBall_j = m_vBPos(j)
            vVel_j = m_vBVelocities(j)
            
            D3DXVec3Subtract vDist, vBall_i, vBall_j  ' A vector linkng the centres of both balls.
            D3DXVec3Subtract vVel_ij, vVel_i, vVel_j        ' Relative velocity between ball "i" and "j".
            If (vDist.x * vDist.x + vDist.z * vDist.z) <= m_2BRadSqrd Then
                ' If the condition above is true, the balls are overlapping. Thus the collision has already occured.
                ' We will pretend that it has happend in this particular point in time, hence:
                Ct = 0
            Else
                ' If the condition is false, then we need to check whether a collision occurs within this time frame.
                ' To do this we use a standard solution for a bionomial equation:
                a = vVel_ij.x * vVel_ij.x + vVel_ij.z * vVel_ij.z
                If a = 0 Then GoTo NextJ
                b = 2 * (vVel_ij.x * vDist.x + vVel_ij.z * vDist.z)
                c = vDist.x * vDist.x + vDist.z * vDist.z - m_2BRadSqrd
                Delta = b * b - 4 * a * c
                If Delta < 0 Then GoTo NextJ
                Ct1 = (-b - Sqr(Delta)) / (2 * a)
                Ct2 = (-b + Sqr(Delta)) / (2 * a)
                ' Take the smaller Ct:
                If Ct1 < Ct2 Then Ct = Ct1 Else Ct = Ct2
            End If
            If Ct >= 0 And Ct <= g_dt Then ' we have a POSSIBLE collision:
                '----------------------
                ' Collision responce:
                '----------------------
                'Frist move both balls along their trajectories to the point of collision:
                D3DXVec3Add vBall_i, vBall_i, ScaleVec3(vVel_i, Ct)
                D3DXVec3Add vBall_j, vBall_j, ScaleVec3(vVel_j, Ct)
                D3DXVec3Subtract vDist, vBall_i, vBall_j
                
                ' Collision's Normal vector:
                D3DXVec3Normalize vCNorm, vDist
                
                ' i'th body velocity at Collision Point:
                vCPVel_i = vVel_i
                ' NOTE: The following three lines of code would by used for computing the vCPVel_i
                '       if this was a "real" 3D collision, but since the balls are always
                '       of the same size, move on the same plane and we don't take friction
                '       into consideration in collision responce, we can say that velocity
                '       at the point of collision is the same as at the ball's centre.
                '
                'D3DXVec3Scale vCP_i, vCNorm, -m_BRadius
                'D3DXVec3Cross vCPVel_i, m_vBSpins(i), vCP_i
                'D3DXVec3Add vCPVel_i, vCPVel_i, vVel_i
                
                ' j'th body velocity at Collision Point:
                vCPVel_j = vVel_j
                'D3DXVec3Scale vCP_j, vCNorm, m_BRadius
                'D3DXVec3Cross vCPVel_j, m_vBSpins(j), vCP_j
                'D3DXVec3Add vCPVel_j, vCPVel_j, vVel_j
                
                ' relative velocity at Collision Point:
                D3DXVec3Subtract vCPVel_ij, vCPVel_i, vCPVel_j
                
                If DotProduct(vCPVel_ij, vCNorm) < 0 Then   ' we have a REAL collision:
                    ' The Collision Impulse value:
                    CMmntm = -(1 + m_BElasticity) * DotProduct(vCPVel_ij, vCNorm) / (2 / m_BMass)
                    ' Update velocities:
                    D3DXVec3Add m_vBVelocities(i), m_vBVelocities(i), ScaleVec3(vCNorm, CMmntm / m_BMass)
                    D3DXVec3Add m_vBVelocities(j), m_vBVelocities(j), ScaleVec3(vCNorm, -CMmntm / m_BMass)
                    ' Update positions:
                    m_vBPos(i) = vBall_i: m_vBPos(j) = vBall_j
                    ' Update the value of momentum for a ball - ball collision,
                    m_CBillBillMomentum = m_CBillBillMomentum + CMmntm
                    ' Update the FirstBallClr variable if neccessary:
                    If j = 0 And m_FirstBallClr = BC_NONE Then
                        m_FirstBallClr = m_PoolBalls.Colour(i)
                    End If
                End If
            End If
NextJ:  Next j
NextI: Next i

End Sub

'--------------------------------------------------------------------------------------------
' Name: CDnR_BallTable()
' Desc: Detects collisions between balls and the table and generates collision responce:
'--------------------------------------------------------------------------------------------
Private Sub CDnR_BallTable()
    Dim vBallPos     As D3DVECTOR    ' Position vector of the i'th ball.
    Dim vBallVel     As D3DVECTOR    ' Velocity vector of a ball.
    Dim D1              As Single       ' A utility variable.
    Dim D2              As Single       ' A utility variable.
    Dim Ct              As Single       ' A point in time when a collision was detected.
    Dim iNorm           As Long         ' Index of a normal vector in the normal vector array
    Dim LatestCt        As Single       ' A point in time, when the most recent collision with an edge took place.
    Dim iLatestEdge     As Single       ' The edge, which took part in the latest collision.
    Dim bCDetected      As Boolean      ' Set to "True" if a collision was detected.
    Dim vCNorm          As D3DVECTOR    ' Collision normal vector.
    Dim vCP_i           As D3DVECTOR    ' Collision Point relative to the centre of the i'th ball
    Dim vCPVel_i        As D3DVECTOR    ' Velocity vector of the i'th ball AT POINT OF COLLISION.
    Dim CMmntm          As Single       ' Collision Momentum.
    Dim vBumperToBill   As D3DVECTOR    ' A vetor linking one vertex of a bumper with a ball's centre point
    Dim MinDistToEdge   As Single       ' Minimal recorded distance to an edge.
    Dim iClosestEdge    As Long         ' Index of the edge, which is at the closest distance to the ball.
    Dim BVelDotTBNorm   As Single       ' Result of a dot product between Ball's velocity vector and a bumper normal.
    
    ' The following variables are used
    ' for checking pocket hits.
    Dim MinDistTime     As Single       ' A point in time, when the distance between a pocket and a ball is in the shortest.
    Dim vMinDist        As D3DVECTOR    ' A vector linking the ball's and the pocket's centre points, when they are at the shortest distance between themselves.
    Dim vDist           As D3DVECTOR    ' A vector linking the ball's and the pocket's centre points.
    Dim BallVelSqrd  As Single       ' Ball's velocity squared.

    
    For i = 0 To m_numPoolBalls - 1
        If Not m_bBInTheGame(i) Then GoTo NextI
        vBallPos = m_vBPos(i)
        vBallVel = m_vBVelocities(i)
        BallVelSqrd = DotProduct(vBallVel, vBallVel)
        ' We are not interested in the balls that are not moving:
        If BallVelSqrd = 0 Then GoTo NextI
        For j = 1 To m_numTableBumpers
            iLatestEdge = 0
            LatestCt = -1
            bCDetected = True
            MinDistToEdge = 10
            iClosestEdge = 0
            
            For k = 1 To 5
                'Computing the distance between the k'th edge of table bumper and the ball:
                D3DXVec3Subtract vBumperToBill, vBallPos, m_TBumpers(j, k)
                D1 = DotProduct(vBumperToBill, m_TBumperNorms(j, k))
                If D1 > m_BRadius Then
                    bCDetected = False
                Else
                    If -D1 < MinDistToEdge Then MinDistToEdge = -D1: iClosestEdge = k
                    BVelDotTBNorm = DotProduct(vBallVel, m_TBumperNorms(j, k))
                    If BVelDotTBNorm = 0 Then GoTo NextK
                    Ct = (m_BRadius - D1) / BVelDotTBNorm
                    If Ct > LatestCt And Ct <= 0 And Ct >= -g_dt Then LatestCt = Ct: iLatestEdge = k
                End If
NextK:      Next k
            If iLatestEdge = 0 Then iNorm = iClosestEdge Else iNorm = iLatestEdge
            If bCDetected Then  ' We have a possible collision:
                vCNorm = m_TBumperNorms(j, iNorm)
                D3DXVec3Scale vCP_i, vCNorm, -m_BRadius
                D3DXVec3Cross vCPVel_i, m_vBSpins(i), vCP_i
                ' i'th ball's velocity at Collision Point:
                D3DXVec3Add vCPVel_i, vCPVel_i, vBallVel
                '----------------------
                ' Collision responce:
                '----------------------
                If DotProduct(vCPVel_i, vCNorm) < 0 Then 'We are sure that we have a collision:
                    ' The Collision Momentum value:
                    CMmntm = -(1 + m_BElasticity) * DotProduct(vCPVel_i, vCNorm) / (1 / m_BMass)
                    If LatestCt > -g_dt Then
                        ' Update ball's position by moving it along its trajectory to the point of collision:
                        D3DXVec3Add vBallPos, vBallPos, ScaleVec3(vBallVel, LatestCt)
                        m_vBPos(i) = vBallPos
                    End If
                    ' Update velocities:
                    BallVelSqrd = DotProduct(vBallVel, vBallVel)
                    D3DXVec3Add vBallVel, vBallVel, ScaleVec3(vCNorm, CMmntm / m_BMass)
                    m_vBVelocities(i) = vBallVel
                    ' Update the value of momentum for a ball - table collision:
                    m_CBillTableMomentum = m_CBillTableMomentum + CMmntm
                End If
            End If
NextJ:  Next j
        ' Now check for a pocket hit:
        ' Compute the point in time, at which the ball will
        ' be at the shortest distance from the pocket and then
        ' check exactly how close will it be.
        For k = 1 To 6
            D3DXVec3Subtract vDist, vBallPos, m_vTPockets(k)
            ' First check if the ball isn't in the pocket already:
            If (vDist.x * vDist.x + vDist.z * vDist.z) <= m_TPocketRadSqrd Then
                ' The i'th ball has fallen into the k'th pocket:
                m_bPocketHitDetected = True
                ' Notify the i'th ball that it has fallen into the k'th pocket:
                m_PoolBalls.FellInPocketNumber(i) = k
                ' Set the ball's position to the centre of the k'th pocket...
                m_vBPos(i) = m_vTPockets(k)
                ' ... and stop its movement.
                m_vBVelocities(i) = vec3(0, 0, 0)
                ' Jump to another ball:
                GoTo NextI
            End If
            ' The second method for finding pocket hits is to compute the point in time,
            ' when the ball is at the shortest distance from the pocket's centre:
            MinDistTime = -DotProduct(vDist, vBallVel) / BallVelSqrd
            ' If the point in time is between the beginning of the time frame (0) and its end (g_dt)...
            If MinDistTime >= 0 And MinDistTime <= g_dt Then
                D3DXVec3Add vMinDist, vDist, ScaleVec3(vBallVel, MinDistTime)
                '... and if the squared module of the minimum distance vector is less or equall to
                ' the squared radius of the pocket, then we have a pocket hit.
                If (vMinDist.x * vMinDist.x + vMinDist.z * vMinDist.z) <= m_TPocketRadSqrd Then
                    ' The i'th ball has fallen into the k'th pocket:
                    m_bPocketHitDetected = True
                    ' Notify the i'th ball that it has fallen into the k'th pocket:
                    m_PoolBalls.FellInPocketNumber(i) = k
                    ' Set the ball's position to the centre of the k'th pocket...
                    m_vBPos(i) = m_vTPockets(k)
                    ' ... and stop its movement.
                    m_vBVelocities(i) = vec3(0, 0, 0)
                    ' Jump to another ball:
                    GoTo NextI
                End If
            End If
        Next k
NextI:
    Next i
End Sub

Friend Sub ClearFirstBallTouchedColour()
    m_FirstBallClr = BC_NONE
End Sub

'-------------------------
' Read-only Properties:
'-------------------------
Friend Property Get BillBillCollisionMomentum() As Single
    BillBillCollisionMomentum = m_CBillBillMomentum
End Property

Friend Property Get BillTableCollisionMomentum() As Single
    BillTableCollisionMomentum = m_CBillTableMomentum
End Property

Friend Property Get PocketHitDetected() As Boolean
    PocketHitDetected = m_bPocketHitDetected
End Property

Friend Property Get FirstBallTouchedColour() As BALL_COLOURS
    FirstBallTouchedColour = m_FirstBallClr
End Property
