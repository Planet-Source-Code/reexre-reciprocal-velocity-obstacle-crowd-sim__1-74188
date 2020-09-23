Attribute VB_Name = "modWORLD"
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
'***********************************************************************************

Option Explicit

Public MaxX        As Single
Public MaxY        As Single
Public Zoom        As Single
Public InvZoom     As Single
Public PanX        As Single
Public PanY        As Single
Public CenX        As Single
Public CenY        As Single
Public Navigating  As Boolean
Public PanZoomChanged As Boolean
Public MaxXPic     As Long
Public MaxYPic     As Long


Public H()         As New clsHuman
Public NH          As Long

Public Const RenderFrame As Long = 3

Public Const VelMULTI As Single = 15

Public CNT         As Long
Public FR As Long

Public Function XtoScreen(X As Single) As Long
    XtoScreen = Zoom * (X - PanX) + CenX
End Function
Public Function YtoScreen(Y As Single) As Long
    YtoScreen = Zoom * (Y - PanY) + CenY
End Function

Public Function xfromScreen(X As Long) As Single
    xfromScreen = (X - CenX) * InvZoom + PanX
End Function
Public Function yfromScreen(Y As Long) As Single
    yfromScreen = (Y - CenY) * InvZoom + PanY
End Function
Public Function IsInsideScreen(X As Long, Y As Long) As Boolean
    ' IsInsideScreen = False

    If X < 0 Then Exit Function
    If X > MaxXPic Then Exit Function
    If Y < 0 Then Exit Function
    If Y > MaxYPic Then Exit Function

    IsInsideScreen = True

End Function
Public Sub DrawGrid(hdc As Long)
    Const C          As Long = 8421504 'RGB(128, 128, 128)
    Dim X0         As Long
    Dim x1         As Long
    Dim x2         As Long
    Dim Y0         As Long
    Dim y1         As Long
    Dim y2         As Long



    Y0 = YtoScreen(0)
    y2 = YtoScreen(MaxY)
    For x1 = 0 To MaxX Step 100
        x2 = XtoScreen(x1 \ 1)
        FastLine hdc, x2, Y0, x2, y2, 2, C
    Next
    X0 = XtoScreen(0)
    x2 = XtoScreen(MaxX)
    For y1 = 0 To MaxY Step 100
        y2 = YtoScreen(y1 \ 1)
        FastLine hdc, X0, y2, x2, y2, 2, C
    Next


    x1 = XtoScreen(0)
    y1 = YtoScreen(0)
    x2 = XtoScreen(MaxX \ 1)
    y2 = YtoScreen(MaxY \ 1)
    FastLine hdc, x1, y1, x2, y1, 3, vbWhite
    FastLine hdc, x1, y2, x2, y2, 3, vbWhite
    FastLine hdc, x1, y1, x1, y2, 3, vbWhite
    FastLine hdc, x2, y1, x2, y2, 3, vbWhite

End Sub

Public Sub allRVO()
    Dim I          As Long
    For I = 1 To NH

        ComputeRVO I, 150, frmMAIN.Check1

    Next
    Separate

End Sub
Public Sub Separate()
    Dim I          As Long
    Dim J          As Long
    Dim Dx         As Single
    Dim Dy         As Single
    Dim MinDist    As Single
    Dim D          As Single
    Dim V          As geoPointVector2D
    Dim Vi         As geoPointVector2D
    Dim Vj         As geoPointVector2D

    Dim VelSum     As Single


    For I = 1 To NH - 1
        For J = I + 1 To NH
            MinDist = H(I).R + H(J).R
            Dx = H(J).X - H(I).X
            If Abs(Dx) < MinDist Then
                Dy = H(J).Y - H(I).Y
                If Abs(Dy) < MinDist Then
                    D = Dx * Dx + Dy * Dy
                    If D < MinDist * MinDist Then

                        D = Sqr(D)
                        D = MinDist - D

                        V = VectorNormalize(mkPoint(Dx, Dy))
                        VelSum = H(I).VEL + H(J).VEL
                        If VelSum > 0 Then
                            Vi = VectorMUL(V, -D * 0.5 * H(I).VEL / VelSum)
                            Vj = VectorMUL(V, D * 0.5 * H(J).VEL / VelSum)
                        Else
                            Vi = VectorMUL(V, -D * 0.5)
                            Vj = VectorMUL(V, D * 0.5)

                        End If

                        H(I).X = H(I).X + Vi.X
                        H(I).Y = H(I).Y + Vi.Y
                        H(J).X = H(J).X + Vj.X
                        H(J).Y = H(J).Y + Vj.Y

                    End If

                End If
            End If

        Next
    Next


End Sub

Public Sub ComputeRVO(a As Long, MaxDist As Single, Optional DebugMode As Boolean = False)
    Dim b          As Long

    Dim Va         As geoPointVector2D
    Dim Vb         As geoPointVector2D

    Dim Pa         As geoPointVector2D
    Dim Pb         As geoPointVector2D

    Dim Vab        As geoPointVector2D

    Dim L1         As geoLine
    Dim L2         As geoLine
    Dim Ca         As geoCircle
    Dim Cb         As geoCircle

    Dim RR         As Single

    Dim Going      As geoPointVector2D

    Dim P1         As geoPointVector2D
    Dim P2         As geoPointVector2D
    Dim D1         As Single
    Dim D2         As Single

    Dim D          As Single
    Dim DMax       As Single

    Dim Dx         As Single
    Dim Dy         As Single

    Dim MustAvoid  As Boolean

    Dim PTS()      As geoPointVector2D
    Dim Npts       As Long

    Dim EYEDist    As Single

    Dim x1         As Long
    Dim y1         As Long
    Dim x2         As Long
    Dim y2         As Long
    Dim X3         As Long
    Dim Y3         As Long

    Dim EyeFromR   As Single
    Dim EyeTOr     As Single
    Dim Perp       As geoPointVector2D

    Dim KK         As Single

    KK = 12

    Pa = mkPoint(H(a).X, H(a).Y)
    Va = mkPoint(H(a).Vx, H(a).Vy)
    Va = VectorMUL(Va, VelMULTI)
    Going = mkPoint(Pa.X + Va.X, Pa.Y + Va.Y)
    If H(a).MaxVEL <> 0 Then
        EYEDist = H(a).R + H(a).VEL * KK * VelMULTI
    End If


    Npts = 0
    DMax = -1
    P1 = AvoidLine(a)
    If P1.X <> -9999 Then
        MustAvoid = True
        Npts = Npts + 1
        ReDim Preserve PTS(Npts)
        PTS(Npts) = P1
    End If
    If MustAvoid Then GoTo SkipPP

    For b = 1 To NH
        If a <> b Then
            Dx = H(b).X - H(a).X
            If Abs(Dx) < EYEDist Then
                Dy = H(b).Y - H(a).Y
                If Abs(Dy) < EYEDist Then

                    'If Abs(AngleDIFF(H(a).ANG, Atan2(Dx, Dy))) < pih Then
                    'Check if human b is in front of a
                    If SameDirPlusMinus90(Va, mkPoint(Dx, Dy)) = 1 Then

                        'Now we build a corridor "field of view" in which the b Human should stay inside
                        '----eye corridor-----------------------------------------------------
                        'Pb is the corridor central farest point
                        'we start from Pa (H(a) pos) and go to Va multplied by KK (Va is alredy
                        'multiplied by VelMULTI)
                        Pb = VectorSUM(Pa, VectorMUL(Va, KK))
                        'Corridor starts with a radius of EyeFromR given by:
                        'EyeFromR = H(a).R + 25 + 10 * (1 - H(a).VEL / H(a).MaxVEL)
                        EyeFromR = H(a).R * 10 '+ 10 * (1 - H(a).VEL / H(a).MaxVEL)
                        'and ends with a radius EyeToR given by:
                        'EyeTOr = H(a).R * 2
                        EyeTOr = H(a).R * 1 * H(a).VEL / H(a).MaxVEL
                        'In Perp we put a vector of length 1 perpendicular to Current Velocioty
                        Perp = VectorNormalize(VectorPERP(Va))
                        'We Build the two lines that describes the corridor
                        L1.P1 = VectorSUM(Pa, VectorMUL(Perp, EyeFromR))
                        L1.P2 = VectorSUM(Pb, VectorMUL(Perp, EyeTOr))
                        L2.P1 = VectorSUB(Pa, VectorMUL(Perp, EyeFromR))
                        L2.P2 = VectorSUB(Pb, VectorMUL(Perp, EyeTOr))
                        If DebugMode And a = 1 Then
                            x1 = XtoScreen(L1.P1.X)
                            y1 = YtoScreen(L1.P1.Y)
                            x2 = XtoScreen(L1.P2.X)
                            y2 = YtoScreen(L1.P2.Y)
                            FastLine frmMAIN.PIC.hdc, x1, y1, x2, y2, 1, vbGreen
                            x1 = XtoScreen(L2.P1.X)
                            y1 = YtoScreen(L2.P1.Y)
                            x2 = XtoScreen(L2.P2.X)
                            y2 = YtoScreen(L2.P2.Y)
                            FastLine frmMAIN.PIC.hdc, x1, y1, x2, y2, 1, vbYellow
                            'x1 = XtoScreen(Going.X)
                            'y1 = YtoScreen(Going.Y)
                            'MyCircle frmMAIN.PIC.hdc, x1, y1, 3, 2, vbRed
                        End If

                        Pb = mkPoint(H(b).X, H(b).Y)
                        'We check if point Pb is inside the corridor
                        'it happens when the point Pb fall at different sides
                        'from lines L1 and L2
                        If Side(L1, Pb) <> Side(L2, Pb) Then
                            'If a = 1 Then
                            '    x1 = XtoScreen(H(a).X)
                            '    y1 = YtoScreen(H(a).Y)
                            '    x2 = XtoScreen(H(b).X)
                            '    y2 = YtoScreen(H(b).Y)
                            '    FastLine frmMAIN.PIC.hdc, x1, y1, x2, y2, 1, vbCyan
                            'End If
                            '--End eyes corridor-----------------------------------------------------------------------------------


                            'Now we build the RVO
                            Vb = mkPoint(H(b).Vx, H(b).Vy)
                            Vb = VectorMUL(Vb, VelMULTI)

                            Ca.Center = Pa
                            Ca.Radius = 0

                            RR = H(a).R + H(b).R + 8
                            ' if pa and pb are very close
                            If (Dx * Dx + Dy * Dy) < RR * RR Then RR = H(b).R:    'Ca.Radius = H(a).R
                            Cb.Center = Pb
                            Cb.Radius = RR

                            'we get 2 main lines of triangle
                            TangentTwoCircles Ca, Cb, L1, L2

                            'Compute (Va+Vb)*.5
                            Vab = VectorMUL(VectorSUM(Va, Vb), 0.5)
                            'And move the triangle to that direction
                            L1.P1 = VectorSUM(L1.P1, Vab)
                            L1.P2 = VectorSUM(L1.P2, Vab)
                            L2.P1 = VectorSUM(L2.P1, Vab)
                            L2.P2 = VectorSUM(L2.P2, Vab)

                            'This is Two Slow
                            'If PointInsideTriangle(Going, L1.P2, L1.P1, L2.P1) Then

                            'We check if "going" is insed the "infinite" triangle
                            'determinated by lines L1 and L2
                            If Side(L1, Going) = -1 Then
                                If Side(L2, Going) = 1 Then
                                    MustAvoid = True
                                    'Now we find the nearest Points on lines
                                    'L1 and L2 to point Going
                                    P1 = NearestFromLine(Going, L1)
                                    P2 = NearestFromLine(Going, L2)
                                    'Of that 2 points we choose the nearest (to going)
                                    D1 = DistFromPointSQ(Going, P1)
                                    D2 = DistFromPointSQ(Going, P2)
                                    If D2 < D1 Then P1 = P2

                                    'Add the nearest point to an array of points PTS
                                    'that will be lately computed
                                    Npts = Npts + 1
                                    ReDim Preserve PTS(Npts)
                                    PTS(Npts) = P1

                                    If DebugMode And a = 1 Then

                                        x1 = XtoScreen(L1.P1.X)
                                        y1 = YtoScreen(L1.P1.Y)
                                        x2 = XtoScreen(L1.P2.X)
                                        y2 = YtoScreen(L1.P2.Y)
                                        X3 = XtoScreen(L2.P1.X)
                                        Y3 = YtoScreen(L2.P1.Y)

                                        FastLine frmMAIN.PIC.hdc, x1, y1, x2, y2, 1, vbRed
                                        FastLine frmMAIN.PIC.hdc, X3, Y3, x2, y2, 1, vbGreen
                                        FastLine frmMAIN.PIC.hdc, X3, Y3, x1, y1, 1, RGB(200, 200, 200)

                                        ''FastLine frmMAIN.PIC.hdc, (Pa.X) \ 1, (Pa.Y) \ 1, (Pa.X + Va.X) \ 1, (Pa.Y + Va.Y) \ 1, 2, vbGreen
                                        ''MyCircle frmMAIN.PIC.hdc, (Pa.X + Va.X) \ 1, (Pa.Y + Va.Y) \ 1, 3, 2, RGB(200, 200, 200)
                                        'X1 = XtoScreen(Going.X)
                                        'Y1 = YtoScreen(Going.Y)
                                        'MyCircle frmMAIN.PIC.hdc, X1, Y1, 3, 2, vbRed
                                        'X1 = XtoScreen(P1.X)
                                        'Y1 = YtoScreen(P1.Y)
                                        'MyCircle frmMAIN.PIC.hdc, X1, Y1, 3, 2, vbCyan

                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

SkipPP:

    If MustAvoid Then
        'We should have computed only 1 point-result for all Overlapping triangles.
        'That means find a polygon of all overlapping traingles and
        'find the point nearest to "going" from all polygon sides.
        'Since at the moment I'm not able to find the result polygon given by
        'a set of n overlapping triangles,
        'We compute someway all the N points nearest to "Going" for every triangle
        'Here there are 3 ways to do that.
        'at the moment I don't know if it's better the (1) or the (3).

        '(1)-Select the wrongets point as the one to compute
        '(could be a good solution)
        '------------------------------------------------------
        'For b = 1 To Npts
        '    D = DistFromPointSQ(Going, PTS(b))
        '    If D > DMax Then DMax = D: P1 = PTS(b)
        'Next

        '------------------------------------------------------
        '(2)-Compute the Average Point
        '(to me do not works very well)
        '------------------------------------------------------
        'P1.X = 0: P1.Y = 0
        'For b = 1 To Npts
        '    P1 = VectorSUM(P1, PTS(b))
        'Next
        'P1 = VectorDIV(P1, CSng(Npts))

        '------------------------------------------------------
        '(3)-Compute the "sum" of
        'all difference between PTS points and "going point"
        '(Seems good) (I have to test more, maybe mode (1) is better
        '------------------------------------------------------
        P1.X = 0: P1.Y = 0
        For b = 1 To Npts
            P1 = VectorSUM(P1, VectorSUB(PTS(b), Going))
        Next
        P1 = VectorSUM(P1, Going)
        '------------------------------------------------------


        P2 = P1
        P1 = VectorSUB(P1, Pa)

        'P1 now contains the results of all RVO for "a-Human"
        'So desired angle to go is
        H(a).DesiredANG = Atan2(P1.X, P1.Y)

        'If P1 do not coincide with Pa
        If VectorMAG(P1) <> 0 Then

            'If DistFromPoint(P2, Pa) < VectorMAG(Va) Then
            '    'If H(a).DesiredVEL > H(a).MaxVEL * 0.05 Then

            'we BRAKE
            H(a).DesiredVEL = H(a).DesiredVEL * (1 - 0.5 * H(a).Shyness) '* 0.99
            
            '    'H(a).DesiredVEL = -H(a).R + DistFromPoint(P2, Pa) / VelMULTI
            '    'If H(a).DesiredVEL < 0 Then H(a).DesiredVEL = 0
            '    'End If
            'Else
            '    H(a).DesiredVEL = H(a).MaxVEL
            'End If
        Else
            H(a).DesiredVEL = H(a).MaxVEL
        End If
    Else
        'Else, If no Human to avoid:
        H(a).DesiredVEL = H(a).MaxVEL
        H(a).DesiredANG = H(a).TargetANG
    End If

    If DebugMode Then If a = 1 Then frmMAIN.PIC.Refresh

End Sub
