Attribute VB_Name = "modGeometry"
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
' Suggestions or new Tools are wellcome!
' Most Function taken from http://paulbourke.net/geometry
'***********************************************************************************

Option Explicit

Public Const PI    As Single = 3.14159265358979    ' Atn (1) * 4
Public Const PI2   As Single = 6.28318530717959    'PI * 2
Public Const PIh   As Single = 1.5707963267949    ' PI * 0.5

Public Type geoPointVector2D
    X              As Single
    Y              As Single
    Bool           As Boolean
End Type

Public Type geoLine
    P1             As geoPointVector2D
    P2             As geoPointVector2D
    ANG            As Single
    Bool           As Boolean
End Type

Public Type geoCircle
    Center         As geoPointVector2D
    Radius         As Single
    Bool           As Boolean
End Type

Public Type geoARC
    Circle         As geoCircle
    A1             As Single
    A2             As Single
    x1             As Single
    y1             As Single
    x2             As Single
    y2             As Single
End Type




Public Function mkPoint(X As Single, Y As Single) As geoPointVector2D
    mkPoint.X = X
    mkPoint.Y = Y
End Function

Public Function mkLine(P1 As geoPointVector2D, P2 As geoPointVector2D) As geoLine
    Dim Dx         As Single
    Dim Dy         As Single

    mkLine.P1 = P1
    mkLine.P2 = P2
    Dx = P2.X - P1.X
    Dy = P2.Y - P1.Y
    mkLine.ANG = Atan2(Dx, Dy)
End Function
Public Sub UpdateLineAng(ByRef L As geoLine)
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = L.P2.X - L.P1.X
    Dy = L.P2.Y - L.P1.Y
    L.ANG = Atan2(Dx, Dy)
    If L.ANG < 0 Then L.ANG = L.ANG + PI2
End Sub
Public Sub UpdateArcPts(ByRef a As geoARC)
    'Knowing A1 and A2 of the arc
    'calc x1,y1 and x2,y2
    With a
        .x1 = .Circle.Center.X + .Circle.Radius * Cos(.A1)
        .y1 = .Circle.Center.Y + .Circle.Radius * Sin(.A1)
        .x2 = .Circle.Center.X + .Circle.Radius * Cos(.A2)
        .y2 = .Circle.Center.Y + .Circle.Radius * Sin(.A2)
    End With
End Sub

Public Function mkLine2(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As geoLine
    Dim Dx         As Single
    Dim Dy         As Single

    mkLine2.P1.X = x1
    mkLine2.P1.Y = y1
    mkLine2.P2.X = x2
    mkLine2.P2.Y = y2
    Dx = x2 - x1
    Dy = y2 - y1
    mkLine2.ANG = Atan2(Dx, Dy)
End Function

Public Function mkCircle(C As geoPointVector2D, R As Single) As geoCircle
    mkCircle.Center = C
    mkCircle.Radius = R
End Function

Public Function mkCircle2(cx As Single, cy As Single, R As Single) As geoCircle
    mkCircle2.Center.X = cx
    mkCircle2.Center.Y = cy
    mkCircle2.Radius = R
End Function

Public Function mkCircle3Points(ByRef P1 As geoPointVector2D, ByRef P2 As geoPointVector2D, ByRef P3 As geoPointVector2D) As geoCircle
    mkCircle3Points.Bool = False

    If privIsNotPerpendicular(P1, P2, P3) Then
        mkCircle3Points = privCircle3Points(P1, P2, P3)
    ElseIf privIsNotPerpendicular(P1, P3, P2) Then
        mkCircle3Points = privCircle3Points(P1, P3, P2)
    ElseIf privIsNotPerpendicular(P2, P1, P3) Then
        mkCircle3Points = privCircle3Points(P2, P1, P3)
    ElseIf privIsNotPerpendicular(P2, P3, P1) Then
        mkCircle3Points = privCircle3Points(P2, P3, P1)
    ElseIf privIsNotPerpendicular(P3, P2, P1) Then
        mkCircle3Points = privCircle3Points(P3, P2, P1)
    ElseIf privIsNotPerpendicular(P3, P1, P2) Then
        mkCircle3Points = privCircle3Points(P3, P1, P2)
    Else
        'msgBox "The three pts are perpendicular to axis"
        If (P2.X - P1.X) = 0 Then
            mkCircle3Points.Center.Y = (P2.Y + P1.Y) / 2
            mkCircle3Points.Center.X = (P3.X + P2.X) / 2
            mkCircle3Points.Radius = DistFromPoint(mkCircle3Points.Center, P2)
            mkCircle3Points.Bool = True
        End If
        If (P3.X - P2.X) = 0 Then
            mkCircle3Points.Center.Y = (P3.Y + P2.Y) / 2
            mkCircle3Points.Center.X = (P2.X + P1.X) / 2
            mkCircle3Points.Radius = DistFromPoint(mkCircle3Points.Center, P2)
            mkCircle3Points.Bool = True
        End If
    End If

End Function

Private Function privCircle3Points(ByRef P1 As geoPointVector2D, ByRef P2 As geoPointVector2D, ByRef P3 As geoPointVector2D) As geoCircle
    Dim aSlope     As Single
    Dim bSlope     As Single
    aSlope = (P2.Y - P1.Y) / (P2.X - P1.X)
    bSlope = (P3.Y - P2.Y) / (P3.X - P2.X)
    If (Abs(aSlope - bSlope) <= 0.000001) Then    'checking whether the given points are colinear.
        MsgBox "The three pts are colinear"
        Exit Function
    End If
    privCircle3Points.Center.X = (aSlope * bSlope * (P1.Y - P3.Y) + bSlope * (P1.X + P2.X) - aSlope * (P2.X + P3.X)) / (2 * (bSlope - aSlope))
    privCircle3Points.Center.Y = -1 * (privCircle3Points.Center.X - (P1.X + P2.X) / 2) / aSlope + (P1.Y + P2.Y) / 2
    privCircle3Points.Radius = DistFromPoint(P1, privCircle3Points.Center)
    privCircle3Points.Bool = True
End Function

Private Function privIsNotPerpendicular(ByRef P1 As geoPointVector2D, ByRef P2 As geoPointVector2D, ByRef P3 As geoPointVector2D) As Boolean
    Dim xDelta_A   As Single
    Dim yDelta_A   As Single
    Dim xDelta_B   As Single
    Dim yDelta_B   As Single

    privIsNotPerpendicular = True

    ' Check the given point are perpendicular to x or y axis
    yDelta_A = P2.Y - P1.Y
    xDelta_A = P2.X - P1.X
    yDelta_B = P3.Y - P2.Y
    xDelta_B = P3.X - P2.X

    ' checking whether the line of the two pts are vertical
    If (Abs(xDelta_A) <= 0.0000001 And Abs(yDelta_B) <= 0.0000001) Then
        'The points are pependicular and parallel to x-y axis
        privIsNotPerpendicular = False
        Exit Function             'return false;
    ElseIf (Abs(yDelta_A) <= 0.0000001) Then
        'A line of two point are perpendicular to x-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(yDelta_B) <= 0.0000001) Then
        'A line of two point are perpendicular to x-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(xDelta_A) <= 0.0000001) Then
        'A line of two point are perpendicular to y-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(xDelta_B) <= 0.0000001) Then
        'A line of two point are perpendicular to y-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    Else

    End If

End Function

Public Function Atan2(X As Single, Y As Single) As Single
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0) * PI
    Else
        Atan2 = -PIh - (Y > 0) * PI
    End If

    If Atan2 < 0 Then Atan2 = Atan2 + PI2

End Function
Public Function FowlerAngle(ByRef Dx As Single, ByRef Dy As Single) As Single
    'Faster than Atan2
    'http://paulbourke.net/geometry/fowler/

    '   This function is due to Rob Fowler.  Given dy and dx between 2 points
    '   A and B, we calculate a number in [0.0, 8.0) which is a monotonic
    '   function of the direction from A to B.
    '
    '   (0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0) correspond to
    '   (  0,  45,  90, 135, 180, 225, 270, 315, 360) degrees, measured
    '   counter-clockwise from the positive x axis.

    Dim Adx        As Single      'Absolute Values of Dx and Dy
    Dim Ady        As Single
    Dim Code       As Long        'Angular Region Classification Code

    Const K = PI2 * 0.125


    Adx = Abs(Dx)                 'Compute the absolute values.
    Ady = Abs(Dy)

    If Adx < Ady Then Code = 1 Else: Code = 0
    If Dx < 0 Then Code = Code + 2
    If Dy < 0 Then Code = Code + 4

    Select Case Code
        Case 0
            If Dx = 0 Then
                FowlerAngle = 0
            Else
                FowlerAngle = Ady / Adx    ';  /* [  0, 45] */
            End If
        Case 1
            FowlerAngle = 2 - (Adx / Ady)    ';      /* ( 45, 90] */
        Case 3
            FowlerAngle = 2 + (Adx / Ady)    ';      /* ( 90,135) */
        Case 2
            FowlerAngle = 4 - (Ady / Adx)    ';      /* [135,180] */
        Case 6
            FowlerAngle = 4 + (Ady / Adx)    ';      /* (180,225] */
        Case 7
            FowlerAngle = 6 - (Adx / Ady)    ';      /* (225,270) */
        Case 5
            FowlerAngle = 6 + (Adx / Ady)    ';      /* [270,315) */
        Case 4
            FowlerAngle = 8 - (Ady / Adx)    ';      /* [315,360) */
    End Select

    FowlerAngle = FowlerAngle * K

End Function
Public Function Atan2Fast1(ByRef X As Single, ByRef Y As Single) As Single
    'http://www.gamedev.net/topic/441464-manually-implementing-atan2-or-atan/
    'maximum error slightly larger than 4 degrees.
    'public single aTan2(single y, single x)
    '{   single coeff_1 = Math.PI / 4d;  single coeff_2 = 3d * coeff_1;  single abs_y = Math.abs(y); single angle;   if (x >= 0d) {      single r = (x - abs_y) / (x + abs_y);       angle = coeff_1 - coeff_1 * r;  } else {        single r = (x + abs_y) / (abs_y - x);       angle = coeff_2 - coeff_1 * r;  }   return y < 0d ? -angle : angle;}
    Const C1       As Single = 0.785398163397448    'atn(1)
    Const C2       As Single = 2.35619449019234    'atn(1)*3
    Dim AbsY       As Single
    Dim R          As Single

    AbsY = Abs(Y)
    If (X >= 0) Then
        R = (X - AbsY) / (X + AbsY)
        Atan2Fast1 = C1 - C1 * R
    Else
        R = (X + AbsY) / (AbsY - X)
        Atan2Fast1 = C2 - C1 * R
    End If

    If Y < 0 Then Atan2Fast1 = -Atan2Fast1

End Function
Public Function Atan2Fast2(ByRef X As Single, ByRef Y As Single) As Single
    'http://lists.apple.com/archives/perfoptimization-dev/2005/Jan/msg00051.html
    '|error| < 0.005 radians

    Dim Z          As Single

    If X = 0 Then
        If (Y > 0) Then Atan2Fast2 = PIh: Exit Function
        If (Y = 0) Then Atan2Fast2 = 0: Exit Function
        Atan2Fast2 = -PIh: Exit Function
    End If

    Z = Y / X
    If (Abs(Z) < 1) Then
        Atan2Fast2 = Z / (1 + 0.28 * Z * Z)
        If (X < 0) Then
            If (Y < 0) Then Atan2Fast2 = Atan2Fast2 + PI: Exit Function
            Atan2Fast2 = Atan2Fast2 + PI: Exit Function
        End If
    Else
        Atan2Fast2 = PIh - Z / (Z * Z + 0.28)
        If (Y < 0) Then Atan2Fast2 = Atan2Fast2 + PI: Exit Function
    End If

    If Atan2Fast2 < 0 Then Atan2Fast2 = Atan2Fast2 + PI2

End Function

Public Function AngleDIFF(A1 As Single, A2 As Single) As Single
    'single difference = secondAngle - firstAngle;

    AngleDIFF = A2 - A1
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend

End Function



Public Function LineLen(ByRef L As geoLine) As Single
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = L.P2.X - L.P1.X
    Dy = L.P2.Y - L.P1.Y
    LineLen = Sqr(Dx * Dx + Dy * Dy)
End Function

Public Function DistFromPoint(ByRef P1 As geoPointVector2D, ByRef P2 As geoPointVector2D) As Single
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = P2.X - P1.X
    Dy = P2.Y - P1.Y
    DistFromPoint = Sqr(Dx * Dx + Dy * Dy)
End Function
Public Function DistFromPointSQ(ByRef P1 As geoPointVector2D, ByRef P2 As geoPointVector2D) As Single
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = P2.X - P1.X
    Dy = P2.Y - P1.Y
    DistFromPointSQ = (Dx * Dx + Dy * Dy)
End Function
Public Function DistFromPoint2(ByRef P As geoPointVector2D, X As Single, Y As Single) As Single
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = X - P.X
    Dy = Y - P.Y
    DistFromPoint2 = Sqr(Dx * Dx + Dy * Dy)
End Function
Public Function DistFromPoint2SQ(ByRef P As geoPointVector2D, X As Single, Y As Single) As Single
    Dim Dx         As Single
    Dim Dy         As Single
    Dx = X - P.X
    Dy = Y - P.Y
    DistFromPoint2SQ = (Dx * Dx + Dy * Dy)
End Function


Public Function DistFromLine(ByRef P As geoPointVector2D, ByRef L As geoLine) As Single
    '
    ' Returns distance from the line, or if the intersecting point on the line nearest
    ' the point tested is outside the endpoints of the line, the distance to the
    ' nearest endpoint.
    '
    ' Returns 9999 on 0 denominator conditions.
    Dim LineMag As Single, u As Single
    Dim iX As Single, iY As Single    ' intersecting point

    LineMag = LineLen(L)
    If LineMag < 0.000001 Then DistFromLine = 9999: Exit Function

    u = (((P.X - L.P1.X) * (L.P2.X - L.P1.X)) + ((P.Y - L.P1.Y) * (L.P2.Y - L.P1.Y)))
    u = u / (LineMag * LineMag)

    'If u < 0.00001 Or u > 1 Then
    '    '// closest point does not fall within the line segment, take the shorter distance
    '    '// to an endpoint
    '    ix = DistFromPoint(P, L.P1)
    '    iy = DistFromPoint(P, L.P2)
    '    If ix > iy Then DistFromLine = iy Else DistFromLine = ix
    'Else
    ' Intersecting point is on the line, use the formula
    iX = L.P1.X + u * (L.P2.X - L.P1.X)
    iY = L.P1.Y + u * (L.P2.Y - L.P1.Y)
    DistFromLine = DistFromPoint2(P, iX, iY)
    'End If

End Function
Public Function DistFromLineSQ(ByRef P As geoPointVector2D, ByRef L As geoLine) As Single
    '
    ' Returns distance from the line, or if the intersecting point on the line nearest
    ' the point tested is outside the endpoints of the line, the distance to the
    ' nearest endpoint.
    '
    ' Returns 9999 on 0 denominator conditions.
    Dim LineMag As Single, u As Single
    Dim iX As Single, iY As Single    ' intersecting point

    LineMag = LineLen(L)
    If LineMag < 0.000001 Then DistFromLineSQ = 999999999: Exit Function

    u = (((P.X - L.P1.X) * (L.P2.X - L.P1.X)) + ((P.Y - L.P1.Y) * (L.P2.Y - L.P1.Y)))
    u = u / (LineMag * LineMag)

    If u < 0.00001 Or u > 1 Then
        '    '// closest point does not fall within the line segment, take the shorter distance
        '    '// to an endpoint
        iX = DistFromPointSQ(P, L.P1)
        iY = DistFromPointSQ(P, L.P2)
        If iX > iY Then DistFromLineSQ = iY Else DistFromLineSQ = iX
    Else
        ' Intersecting point is on the line, use the formula
        iX = L.P1.X + u * (L.P2.X - L.P1.X)
        iY = L.P1.Y + u * (L.P2.Y - L.P1.Y)
        DistFromLineSQ = DistFromPoint2SQ(P, iX, iY)
    End If

End Function
Public Function NearestFromLine(ByRef P As geoPointVector2D, ByRef L As geoLine) As geoPointVector2D
    '
    ' Returns distance from the line, or if the intersecting point on the line nearest
    ' the point tested is outside the endpoints of the line, the distance to the
    ' nearest endpoint.
    '
    ' Returns 9999 on 0 denominator conditions.
    Dim LineMag As Single, u As Single
    Dim iX As Single, iY As Single    ' intersecting point
    NearestFromLine.Bool = False
    LineMag = LineLen(L)
    If LineMag < 0.000001 Then Exit Function

    u = (((P.X - L.P1.X) * (L.P2.X - L.P1.X)) + ((P.Y - L.P1.Y) * (L.P2.Y - L.P1.Y)))
    u = u / (LineMag * LineMag)

    NearestFromLine.Bool = True
    If u < 0.00001 Or u > 1 Then
        '// closest point does not fall within the line segment, take the shorter distance
        '// to an endpoint
        iX = DistFromPoint(P, L.P1)
        iY = DistFromPoint(P, L.P2)
        If iX < iY Then NearestFromLine = L.P1 Else NearestFromLine = L.P2
    Else
        ' Intersecting point is on the line, use the formula
        NearestFromLine.X = L.P1.X + u * (L.P2.X - L.P1.X)
        NearestFromLine.Y = L.P1.Y + u * (L.P2.Y - L.P1.Y)

    End If

End Function
Public Function IntersectOfLines(ByRef L1 As geoLine, ByRef L2 As geoLine) As geoPointVector2D
    Dim D          As Single
    Dim NA         As Single
    Dim NB         As Single
    Dim DX1        As Single
    Dim DX2        As Single
    Dim DY1        As Single
    Dim DY2        As Single
    Dim uA         As Single
    Dim uB         As Single

    DX1 = L1.P2.X - L1.P1.X
    DY1 = L1.P2.Y - L1.P1.Y
    DX2 = L2.P2.X - L2.P1.X
    DY2 = L2.P2.Y - L2.P1.Y

    ' Denominator for ua and ub are the same, so store this calculation
    D = (DY2) * (DX1) - _
        (DX2) * (DY1)

    'NA and NB are calculated as seperate values for readability
    NA = (DX2) * (L1.P1.Y - L2.P1.Y) - _
         (DY2) * (L1.P1.X - L2.P1.X)

    NB = (DX1) * (L1.P1.Y - L2.P1.Y) - _
         (DY1) * (L1.P1.X - L2.P1.X)

    ' Make sure there is not a division by zero - this also indicates that
    ' the lines are parallel.
    ' If NA and NB were both equal to zero the lines would be on top of each
    ' other (coincidental).  This check is not done because it is not
    ' necessary for this implementation (the parallel check accounts for this).
    IntersectOfLines.Bool = False

    If D = 0 Then Exit Function

    ' Calculate the intermediate fractional point that the lines potentially intersect.
    uA = NA / D

    ' The fractional point will be between 0 and 1 inclusive if the lines
    ' intersect.  If the fractional calculation is larger than 1 or smaller
    ' than 0 the lines would need to be longer to intersect.
    If uA >= 0 Then
        If uA <= 1 Then
            ' Calculate the intermediate fractional point that the lines potentially intersect.
            uB = NB / D
            If uB >= 0 Then
                If uB <= 1 Then
                    IntersectOfLines.X = L1.P1.X + (uA * (DX1))
                    IntersectOfLines.Y = L1.P1.Y + (uA * (DY1))
                    IntersectOfLines.Bool = True
                End If
            End If
        End If
    End If

End Function
Public Function IntersectOfLines2(ByRef L1 As geoLine, ByRef L2 As geoLine) As geoPointVector2D
    '********************************************
    '*  Intersection of LINES (not segments)    *
    '********************************************

    Dim D          As Single
    Dim NA         As Single
    Dim NB         As Single
    Dim DX1        As Single
    Dim DX2        As Single
    Dim DY1        As Single
    Dim DY2        As Single
    Dim uA         As Single
    Dim uB         As Single

    DX1 = L1.P2.X - L1.P1.X
    DY1 = L1.P2.Y - L1.P1.Y
    DX2 = L2.P2.X - L2.P1.X
    DY2 = L2.P2.Y - L2.P1.Y

    ' Denominator for ua and ub are the same, so store this calculation
    D = (DY2) * (DX1) - _
        (DX2) * (DY1)

    'NA and NB are calculated as seperate values for readability
    NA = (DX2) * (L1.P1.Y - L2.P1.Y) - _
         (DY2) * (L1.P1.X - L2.P1.X)

    NB = (DX1) * (L1.P1.Y - L2.P1.Y) - _
         (DY1) * (L1.P1.X - L2.P1.X)

    ' Make sure there is not a division by zero - this also indicates that
    ' the lines are parallel.
    ' If NA and NB were both equal to zero the lines would be on top of each
    ' other (coincidental).  This check is not done because it is not
    ' necessary for this implementation (the parallel check accounts for this).
    IntersectOfLines2.Bool = False

    If D = 0 Then Exit Function

    ' Calculate the intermediate fractional point that the lines potentially intersect.
    uA = NA / D
    ' Calculate the intermediate fractional point that the lines potentially intersect.
    'uB = NB / D

    IntersectOfLines2.X = L1.P1.X + (uA * (DX1))
    IntersectOfLines2.Y = L1.P1.Y + (uA * (DY1))
    IntersectOfLines2.Bool = True


End Function
Public Sub IntersectCircleLine(ByRef C As geoCircle, _
                               ByRef L As geoLine, _
                               ByRef Sol1 As geoPointVector2D, _
                               ByRef Sol2 As geoPointVector2D)


    Dim Dx         As Single
    Dim Dy         As Single
    Dim I          As Single
    Dim AA         As Single
    Dim BB         As Single
    Dim CC         As Single
    Dim mu         As Single

    Sol1.Bool = False
    Sol2.Bool = False


    Dx = L.P2.X - L.P1.X
    Dy = L.P2.Y - L.P1.Y

    AA = Dx * Dx + Dy * Dy        '
    If AA = 0 Then Exit Sub

    BB = 2 * ((Dx) * (L.P1.X - C.Center.X) + _
              (Dy) * (L.P1.Y - C.Center.Y))


    CC = (C.Center.X) ^ 2 + (C.Center.Y) ^ 2 + _
         (L.P1.X) ^ 2 + _
         (L.P1.Y) ^ 2 - _
         2 * (C.Center.X * L.P1.X + C.Center.Y * L.P1.Y) - (C.Radius) ^ 2

    I = BB * BB - 4 * AA * CC


    Select Case I
        Case Is < 0
            'No intersection
            Exit Sub
        Case 0
            'one intersection
            Sol1.Bool = True
            mu = -BB / (2 * AA)
            Sol1.X = L.P1.X + mu * (Dx)
            Sol1.Y = L.P1.Y + mu * (Dy)
        Case Is > 0
            ' two intersections
            ' first intersection
            Sol1.Bool = True
            Sol2.Bool = True
            mu = (-BB + Sqr(BB * BB - 4 * AA * CC)) / (2 * AA)
            Sol1.X = L.P1.X + mu * (Dx)
            Sol1.Y = L.P1.Y + mu * (Dy)
            ' second intersection
            mu = (-BB - Sqr(BB * BB - 4 * AA * CC)) / (2 * AA)
            Sol2.X = L.P1.X + mu * (Dx)
            Sol2.Y = L.P1.Y + mu * (Dy)

    End Select

    'to make this work for "LINE SEGMENT"
    If NearestFromLine(Sol1, L).Bool = False Then Sol1.Bool = False
    If NearestFromLine(Sol2, L).Bool = False Then Sol2.Bool = False


End Sub

Public Sub IntersectOfCircles(ByRef C1 As geoCircle, _
                              ByRef C2 As geoCircle, _
                              ByRef Sol1 As geoPointVector2D, _
                              ByRef Sol2 As geoPointVector2D)

    Dim D          As Single
    Dim c1R        As Single
    Dim c2R        As Single
    Dim M          As Single
    Dim N          As Single
    Dim a          As Single
    Dim H          As Single
    Dim P          As geoPointVector2D

    'Calculate distance between centres of circle
    D = DistFromPoint(C1.Center, C2.Center)
    c1R = C1.Radius
    c2R = C2.Radius
    M = c1R + c2R
    N = c1R - c2R
    If (N < 0) Then N = -N

    Sol1.Bool = False
    Sol2.Bool = False

    'No solns
    If (D > M) Then Exit Sub

    'Circle are contained within each other
    If (D < N) Then Exit Sub

    'Circles are the same
    If (D = 0) And (c1R = c2R) Then Exit Sub

    'Solve for a
    a = (c1R * c1R - c2R * c2R + D * D) / (2 * D)

    'Solve for h
    H = Sqr(c1R * c1R - a * a)

    'Calculate point p, where the line through the circle intersection points crosses the line between the circle centers.
    P.X = C1.Center.X + (a / D) * (C2.Center.X - C1.Center.X)
    P.Y = C1.Center.Y + (a / D) * (C2.Center.Y - C1.Center.Y)

    '1 soln , circles are touching
    If D = (c1R + c2R) Then Sol1 = P: Sol1.Bool = True: Exit Sub

    '2solns
    H = H / D
    Sol1.X = P.X + (H) * (C2.Center.Y - C1.Center.Y)
    Sol1.Y = P.Y - (H) * (C2.Center.X - C1.Center.X)
    Sol2.X = P.X - (H) * (C2.Center.Y - C1.Center.Y)
    Sol2.Y = P.Y + (H) * (C2.Center.X - C1.Center.X)
    Sol1.Bool = True
    Sol2.Bool = True

End Sub

Public Function VectorProject(ByRef V As geoPointVector2D, ByRef Vto As geoPointVector2D) As geoPointVector2D
    'Poject Vector V to vector Vto
    Dim K          As Single
    Dim D          As Single

    D = Sqr(Vto.X * Vto.X + Vto.Y * Vto.Y)
    If D = 0 Then Exit Function
    K = (V.X * Vto.X + V.Y * Vto.Y) / D

    VectorProject.X = (Vto.X / D) * K
    VectorProject.Y = (Vto.Y / D) * K

End Function

Public Function VectorReflect(ByRef V As geoPointVector2D, ByRef wall As geoPointVector2D) As geoPointVector2D
    'Function returning the reflection of one vector around another.
    'it's used to calculate the rebound of a Vector on another Vector
    'Vector "V" represents current velocity of a point.
    'Vector "Wall" represent the angle of a wall where the point Bounces.
    'Returns the vector velocity that the point takes after the rebound

    Dim vDot       As Single
    Dim D          As Single
    Dim NwX        As Single
    Dim NwY        As Single

    D = Sqr(wall.X * wall.X + wall.Y * wall.Y)
    If D = 0 Then Exit Function

    NwX = wall.X / D
    NwY = wall.Y / D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.X * NwX + V.Y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.X = -V.X + NwX
    VectorReflect.Y = -V.Y + NwY


End Function
Public Function VectorSUM(ByRef V1 As geoPointVector2D, ByRef V2 As geoPointVector2D) As geoPointVector2D
    VectorSUM.X = V1.X + V2.X
    VectorSUM.Y = V1.Y + V2.Y
End Function
Public Function VectorSUB(ByRef V1 As geoPointVector2D, ByRef V2 As geoPointVector2D) As geoPointVector2D
    VectorSUB.X = V1.X - V2.X
    VectorSUB.Y = V1.Y - V2.Y
End Function
Public Function VectorMUL(ByRef V As geoPointVector2D, ByRef Value As Single) As geoPointVector2D
    'Scalar
    VectorMUL.X = V.X * Value
    VectorMUL.Y = V.Y * Value
End Function
Public Function VectorDIV(ByRef V As geoPointVector2D, ByRef Value As Single) As geoPointVector2D
    'Scalar
    If Value = 0 Then Exit Function
    VectorDIV.X = V.X / Value
    VectorDIV.Y = V.Y / Value
End Function
Public Function VectorDOT(ByRef V1 As geoPointVector2D, ByRef V2 As geoPointVector2D) As Single
    'Dot Product
    VectorDOT = V1.X * V2.X + V1.Y * V2.Y
End Function
Public Function VectorCROSS(ByRef V1 As geoPointVector2D, ByRef V2 As geoPointVector2D) As Single
    'The cross product is a 3D thing, which do not have sense in a 2D world.
    VectorCROSS = V1.X * V2.Y - V2.X * V1.Y
End Function

Public Function VectorMAG(ByRef V As geoPointVector2D) As Single
    'V magnitude
    VectorMAG = Sqr(V.X * V.X + V.Y * V.Y)
End Function
Public Function VectorNormalize(ByRef V As geoPointVector2D) As geoPointVector2D
    'convert vector to UNIT length
    Dim M          As Single
    M = VectorMAG(V)
    If M = 0 Then Exit Function
    VectorNormalize.X = V.X / M
    VectorNormalize.Y = V.Y / M
    'VectorNormalize = VectorDIV(V, M)
End Function
Public Function VectorNormal(ByRef V As geoPointVector2D) As geoPointVector2D
    'Normal [Perpendicular]
    VectorNormal.X = -V.Y
    VectorNormal.Y = V.X
End Function
Public Function VectorPERP(ByRef V As geoPointVector2D) As geoPointVector2D
    'Normal [Perpendicular]
    VectorPERP.X = -V.Y
    VectorPERP.Y = V.X
End Function


Public Sub TangentTwoCircles(ByRef C1 As geoCircle, ByRef C2 As geoCircle, _
                             ByRef retL1 As geoLine, ByRef retL2 As geoLine)
    'by Roberto Mior (reexre)

    Dim C3         As geoCircle
    Dim R3         As Single
    Dim CM         As geoCircle

    Dim L1P1       As geoPointVector2D
    Dim L1P2       As geoPointVector2D
    Dim L2P1       As geoPointVector2D
    Dim L2P2       As geoPointVector2D

    Dim A1         As Single
    Dim A2         As Single
    Dim Offset     As Single

    CM.Center.X = (C1.Center.X + C2.Center.X) * 0.5
    CM.Center.Y = (C1.Center.Y + C2.Center.Y) * 0.5
    CM.Radius = DistFromPoint(C1.Center, C2.Center) * 0.5

    R3 = C1.Radius - C2.Radius
    If R3 > 0 Then
        C3.Center = C1.Center
        C3.Radius = R3
        L1P2 = C2.Center
        L2P2 = C2.Center
        Offset = C2.Radius
    Else
        C3.Center = C2.Center
        C3.Radius = -R3
        L1P2 = C1.Center
        L2P2 = C1.Center
        Offset = C1.Radius
    End If

    IntersectOfCircles CM, C3, L1P1, L2P1

    If L1P1.Bool Or L1P2.Bool Then

        retL1 = mkLine(L1P1, L1P2)
        retL2 = mkLine(L2P1, L2P2)

        A1 = retL1.ANG + PIh
        A2 = retL2.ANG - PIh

        retL1.P1.X = retL1.P1.X + Cos(A1) * Offset
        retL1.P2.X = retL1.P2.X + Cos(A1) * Offset
        retL1.P1.Y = retL1.P1.Y + Sin(A1) * Offset
        retL1.P2.Y = retL1.P2.Y + Sin(A1) * Offset

        retL2.P1.X = retL2.P1.X + Cos(A2) * Offset
        retL2.P2.X = retL2.P2.X + Cos(A2) * Offset
        retL2.P1.Y = retL2.P1.Y + Sin(A2) * Offset
        retL2.P2.Y = retL2.P2.Y + Sin(A2) * Offset

    End If

End Sub





Public Function LineOffset(L As geoLine, D As Single, Optional LeftSide As Boolean = False) As geoLine
    Dim iX         As Single
    Dim iY         As Single
    Dim S          As Single

    UpdateLineAng L

    S = IIf(LeftSide, -1, 1)

    iX = S * D * Cos(L.ANG + PIh)
    iY = S * D * Sin(L.ANG + PIh)

    LineOffset.P1.X = L.P1.X + iX
    LineOffset.P1.Y = L.P1.Y + iY
    LineOffset.P2.X = L.P2.X + iX
    LineOffset.P2.Y = L.P2.Y + iY


End Function


Public Function Fillet(ByRef L1 As geoLine, ByRef L2 As geoLine, Radius As Single, retArc As geoARC, Optional ModifyLines As Boolean = False)
    'by Roberto Mior (reexre)
    'Find Arc (of a given radius) tangent to two lines

    Dim tmpL1      As geoLine
    Dim tmpL2      As geoLine
    Dim P          As geoPointVector2D
    Dim IntesectP  As geoPointVector2D
    Dim ArcCenterP As geoPointVector2D
    Dim I          As Long
    Dim J          As Long

    Dim L(1 To 4)  As geoLine

    Dim A1         As Single
    Dim A2         As Single
    Dim A3         As Single

    Dim arcP1      As geoPointVector2D
    Dim arcP2      As geoPointVector2D


    IntesectP = IntersectOfLines2(L1, L2)

    If DistFromPoint(IntesectP, L1.P1) < DistFromPoint(IntesectP, L1.P2) Then
        tmpL1.P1 = IntesectP
        tmpL1.P2 = L1.P2
    Else
        tmpL1.P1 = L1.P1
        tmpL1.P2 = IntesectP
    End If

    If DistFromPoint(IntesectP, L2.P1) < DistFromPoint(IntesectP, L2.P2) Then
        tmpL2.P1 = IntesectP
        tmpL2.P2 = L2.P2
    Else
        tmpL2.P1 = L2.P1
        tmpL2.P2 = IntesectP
    End If


    L(1) = LineOffset(tmpL1, Radius, False)
    L(2) = LineOffset(tmpL1, Radius, True)
    L(3) = LineOffset(tmpL2, Radius, False)
    L(4) = LineOffset(tmpL2, Radius, True)


    ' Find intersection point of 4 offset lines
    ArcCenterP.Bool = False
    For I = 1 To 3
        For J = I + 1 To 4
            Debug.Print I, J
            P = IntersectOfLines(L(I), L(J))
            If P.Bool Then ArcCenterP = P
            If ArcCenterP.Bool Then: Exit For
        Next
        If ArcCenterP.Bool Then: Exit For
    Next
    '------------------------------------------

    retArc.Circle.Center = ArcCenterP
    retArc.Circle.Radius = Radius

    arcP1 = NearestFromLine(ArcCenterP, tmpL1)
    arcP2 = NearestFromLine(ArcCenterP, tmpL2)

    If arcP1.Bool = False Then ArcCenterP.Bool = False
    If arcP2.Bool = False Then ArcCenterP.Bool = False

    'conpute arc "start" and "end" angles
    A1 = Atan2(arcP1.X - ArcCenterP.X, arcP1.Y - ArcCenterP.Y)
    A2 = Atan2(arcP2.X - ArcCenterP.X, arcP2.Y - ArcCenterP.Y)

    If AngleDIFF(A1, A2) > 0 Then
        retArc.A1 = A1
        retArc.A2 = A2
    Else
        retArc.A1 = A2
        retArc.A2 = A1
    End If

    If ModifyLines And ArcCenterP.Bool Then
        If DistFromPoint(arcP1, tmpL1.P1) < DistFromPoint(arcP1, tmpL1.P2) Then
            L1.P1 = arcP1
        Else
            L1.P2 = arcP1
        End If
        If DistFromPoint(arcP2, L2.P1) < DistFromPoint(arcP2, L2.P2) Then
            L2.P1 = arcP2
        Else
            L2.P2 = arcP2
        End If
    End If

    UpdateArcPts retArc

End Function

Public Function PointInsideTriangle(P As geoPointVector2D, a As geoPointVector2D, b As geoPointVector2D, C As geoPointVector2D) As Boolean
    'http://www.blackpawn.com/texts/pointinpoly/default.html
    'with barycentric coordinates
    Dim V0         As geoPointVector2D
    Dim V1         As geoPointVector2D
    Dim V2         As geoPointVector2D
    Dim dot00      As Single
    Dim dot01      As Single
    Dim dot02      As Single
    Dim dot11      As Single
    Dim dot12      As Single
    Dim InvDenom   As Single
    Dim u          As Single
    Dim V          As Single

    '// Compute vectors
    'V0 = C - A
    'V1 = B - A
    'V2 = P - A
    V0 = VectorSUB(C, a)
    V1 = VectorSUB(b, a)
    V2 = VectorSUB(P, a)

    '// Compute dot products
    'dot00 = dot(V0, V0)
    'dot01 = dot(V0, V1)
    'dot02 = dot(V0, V2)
    'dot11 = dot(V1, V1)
    'dot12 = dot(V1, V2)

    dot00 = VectorDOT(V0, V0)
    dot01 = VectorDOT(V0, V1)
    dot02 = VectorDOT(V0, V2)
    dot11 = VectorDOT(V1, V1)
    dot12 = VectorDOT(V1, V2)


    '// Compute barycentric coordinates
    InvDenom = (dot00 * dot11 - dot01 * dot01)
    If InvDenom = 0 Then Exit Function
    InvDenom = 1 / InvDenom
    u = (dot11 * dot02 - dot01 * dot12) * InvDenom
    V = (dot00 * dot12 - dot01 * dot02) * InvDenom

    '// Check if point is in triangle
    'return (u >= 0) && (v >= 0) && (u + v < 1)
    PointInsideTriangle = (u >= 0) And (V >= 0) And ((u + V) < 1)
End Function
Public Function PointInsideTriangle2(L1 As geoLine, L2 As geoLine, L3 As geoLine, P As geoPointVector2D) As Boolean
    'not tested yet
    Dim S1         As Long
    Dim S2         As Long
    Dim S3         As Long
    S1 = Side(L1, P)
    S2 = Side(L2, P)
    If S1 = S2 Then PointInsideTriangle2 = (S1 = Side(L3, P))
End Function

Public Function Side(L As geoLine, P As geoPointVector2D) As Long
    'Returns which side of a line is a point
    Dim vec        As geoPointVector2D
    Dim D          As Single

    vec = VectorSUB(L.P2, L.P1)
    vec = VectorPERP(vec)
    D = VectorDOT(vec, VectorSUB(P, L.P1))
    Side = Sgn(D)
End Function

Public Function SameDirPlusMinus90(P1 As geoPointVector2D, P2 As geoPointVector2D) As Long
    'similar to "side" but + - 90°
    SameDirPlusMinus90 = Sgn(VectorDOT(P1, P2))
End Function
