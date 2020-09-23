Attribute VB_Name = "modLineOBS"
Option Explicit

Public NL          As Long
Public L()         As geoLine


Public Sub AddLine(x1, y1, x2, y2)
    NL = NL + 1
    ReDim Preserve L(NL)
    With L(NL)
        .P1.X = x1
        .P1.Y = y1
        .P2.X = x2
        .P2.Y = y2

    End With
    UpdateLineAng L(NL)

End Sub

Public Sub DrawLines(hdc As Long)
    Dim I          As Long
    For I = 1 To NL
        With L(I)
            FastLine hdc, .P1.X \ 1, .P1.Y \ 1, .P2.X \ 1, .P2.Y \ 1, 1, vbWhite
        End With
    Next

End Sub

Public Function AvoidLine(I As Long) As geoPointVector2D


    Dim D          As Single
    Dim D1         As Single
    Dim D2         As Single

    Dim J          As Long
    Dim HIpos      As geoPointVector2D
    Dim P          As geoPointVector2D
    Dim P1         As geoPointVector2D
    Dim P2         As geoPointVector2D

    Dim C1         As geoCircle
    Dim C2         As geoCircle
    Dim L1         As geoLine
    Dim L2         As geoLine

    Dim LA         As geoLine

    Dim Going      As geoPointVector2D


    AvoidLine.X = -9999
    Exit Function

    HIpos = mkPoint(H(I).X, H(I).Y)
    Going.X = HIpos.X + H(I).Vx * VelMULTI
    Going.Y = HIpos.Y + H(I).Vy * VelMULTI




    For J = 1 To NL
        D = DistFromLineSQ(HIpos, L(J))
        If D < 40000 Then
            If D < H(I).R * H(I).R Then
                '                Stop

                'separate
                P1 = NearestFromLine(HIpos, L(J))
                P1 = VectorSUB(HIpos, P1)
                P1 = VectorNormalize(P1)
                P1 = VectorMUL(P1, H(I).R - Sqr(D))
                P1 = VectorSUM(HIpos, P1)
                H(I).X = P1.X
                H(I).Y = P1.Y

            Else


                LA = L(J)
                P = mkPoint((2 * H(I).R) * Cos(LA.ANG), (2 * H(I).R) * Sin(LA.ANG))
                LA.P1 = VectorSUB(LA.P1, P)
                LA.P2 = VectorSUM(LA.P2, P)


                L1.P1 = HIpos
                L2.P1 = HIpos
                L1.P2 = LA.P1
                L2.P2 = LA.P2


                'L1.P1.X = L1.P1.X + H(I).Vx
                'L1.P1.Y = L1.P1.Y + H(I).Vy
                'L1.P2.X = L1.P2.X + H(I).Vx
                'L1.P2.Y = L1.P2.Y + H(I).Vy
                'If L1.P1.Bool Or L1.P2.Bool Then
                FastLine frmMAIN.PIC.hdc, L1.P1.X \ 1, L1.P1.Y \ 1, _
                         L1.P2.X \ 1, L1.P2.Y \ 1, 1, vbRed
                FastLine frmMAIN.PIC.hdc, L2.P1.X \ 1, L2.P1.Y \ 1, _
                         L2.P2.X \ 1, L2.P2.Y \ 1, 1, vbGreen




                If Side(L2, Going) <> Side(L1, Going) Then

                    '                        Stop
                    P1 = NearestFromLine(Going, L1)
                    P2 = NearestFromLine(Going, L2)
                    D1 = DistFromPointSQ(Going, P1)
                    D2 = DistFromPointSQ(Going, P2)
                    If D2 < D1 Then P1 = P2
                    AvoidLine = P1


                End If
            End If
        End If

    Next

End Function
