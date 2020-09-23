VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
'***********************************************************************************

Option Explicit

Public X           As Single
Public Y           As Single
Public TX          As Single
Public TY          As Single

Public Vx          As Single      'Dont change these values, change angle and vel instead
Public Vy          As Single

Public VEL         As Single
Public DesiredVEL  As Single
Public MaxVEL      As Single

Public ANG         As Single
Private CosA       As Single
Private SinA       As Single

Public DesiredANG  As Single
Public TargetANG   As Single

Public Shyness     As Single


Public R           As Single

Public MyColor     As Long

Private Trail()    As geoPointVector2D
Private ScreenTrailX() As Long
Private ScreenTrailY() As Long

Public Ntrail      As Long
Private ReachTime  As Long

Private DistDone   As Single


Public Sub InitTrail(N As Long)
    Ntrail = N
    ReDim Trail(N)
    ReDim ScreenTrailX(N)
    ReDim ScreenTrailY(N)

End Sub

'Human speed= 1.33 meters per second
Public Sub RenderTrail(hdc As Long)
    
    Dim I          As Long
    Dim x1         As Long
    Dim y1         As Long
    Dim x2         As Long
    Dim y2         As Long
   

    Dim II         As Long

    Dim Nt2        As Long

    Const W As Long = 2

    x1 = XtoScreen(X)
    y1 = YtoScreen(Y)
    If Not (IsInsideScreen(x1, y1)) Then Exit Sub
    ' DRAW TRAIL---------------------------------------------------
    '    For I = 1 To Ntrail
    '        x1 = XtoScreen(Trail(I).X)
    '        y1 = YtoScreen(Trail(I).Y)
    '        SetPixel hdc, x1, y1, MyColor
    '    Next

    For I = 0 To Ntrail - 1
        ScreenTrailX(I) = XtoScreen(Trail(I).X)
        ScreenTrailY(I) = YtoScreen(Trail(I).Y)
    Next

    II = ((CNT \ 5) Mod Ntrail)
    For I = 1 To II
        FastLine hdc, ScreenTrailX(I - 1), ScreenTrailY(I - 1), ScreenTrailX(I), ScreenTrailY(I), W, MyColor
    Next
    II = II + 1
    Nt2 = Ntrail - 2
    For I = II To Nt2
        FastLine hdc, ScreenTrailX(I + 1), ScreenTrailY(I + 1), ScreenTrailX(I), ScreenTrailY(I), W, MyColor
    Next
    If II <> Ntrail Then
        I = Ntrail - 1            ':MyCircle hdc, ScreenTrailX(I), ScreenTrailY(I), 5, 1, vbCyan
        II = 0                    ':MyCircle hdc, ScreenTrailX(II), ScreenTrailY(II), 5, 1, vbRed
        FastLine hdc, ScreenTrailX(I), ScreenTrailY(I), ScreenTrailX(II), ScreenTrailY(II), W, MyColor
    End If
    '-----------------------------------------------------------------
    
End Sub
Public Sub Render(hdc As Long)
    Dim I          As Long
    Dim x1         As Long
    Dim y1         As Long
    Dim x2         As Long
    Dim y2         As Long
    Dim sX1        As Single
    Dim sY1        As Single
    Dim sX2        As Single
    Dim sY2        As Single

    Dim Rh         As Long


    
    Rh = (R * 0.5) * Zoom

    x1 = XtoScreen(X)
    y1 = YtoScreen(Y)
    If Not (IsInsideScreen(x1, y1)) Then Exit Sub
    MyCircle hdc, x1, y1, Rh, R * Zoom, MyColor
    'MyCircle hdc, x1, y1, R * Zoom, 1, vbYellow

    x2 = XtoScreen(X + CosA * R + Vx * 2)
    y2 = YtoScreen(Y + SinA * R + Vy * 2)
    FastLine hdc, x1, y1, x2, y2, 2, vbYellow    ' MyColor

    '--------------------------------------draw legs
    'sX1 = X - SinA
    'sY1 = Y + CosA
    'sX2 = sX1 + CosA * Cos(DistDone * 0.5) * VEL * 5
    'sY2 = sY1 + SinA * Cos(DistDone * 0.5) * VEL * 5
    'x1 = XtoScreen(sX1)
    'y1 = YtoScreen(sY1)
    'x2 = XtoScreen(sX2)
    'y2 = YtoScreen(sY2)
    'FastLine hdc, x1, y1, x2, y2, 1, vbRed
    'sX1 = X + SinA
    'sY1 = Y - CosA
    'sX2 = sX1 - CosA * Cos(DistDone * 0.5) * VEL * 5
    'sY2 = sY1 - SinA * Cos(DistDone * 0.5) * VEL * 5
    'x2 = XtoScreen(sX1)
    'y2 = YtoScreen(sY1)
    'FastLine hdc, x1, y1, x2, y2, 1, MyColor
    'x1 = XtoScreen(sX2)
    'y1 = YtoScreen(sY2)
    'FastLine hdc, x1, y1, x2, y2, 1, vbRed
    '---------------------------------------------



    x1 = XtoScreen(TX)
    y1 = YtoScreen(TY)
    MyCircle hdc, x1, y1, 2, R * 0.2 * Zoom, MyColor



End Sub

Public Sub MOVE(ByRef MyIDX As Long)
    Dim Dx         As Single
    Dim Dy         As Single

    Dim AD         As Single

    If CNT Mod 5 = 0 Then
        If Ntrail > 0 Then
            Trail((CNT \ 5) Mod Ntrail).X = X
            Trail((CNT \ 5) Mod Ntrail).Y = Y
        End If
    End If

    If CNT < ReachTime + 200 Then Vx = Vx * 0.9: Vy = Vy * 0.9: Exit Sub

    AD = AngleDIFF(DesiredANG, ANG)

    '    ANG = ANG - 0.01 * Sgn(AngleDIFF(DesiredANG, ANG))
    '    ANG = ANG - (0.005 + Shyness * 0.005) * Sgn(AngleDIFF(DesiredANG, ANG))
    'ANG = ANG - 0.02 * (1 + Shyness) * (AD + Sgn(AD) * 0.2)
    
    ANG = ANG - 0.05 * (1 + Shyness) * (AD + Sgn(AD) * 0.01)
    
    VEL = VEL + (DesiredVEL - VEL) * 0.01


    DistDone = DistDone + VEL

    CosA = Cos(ANG)
    SinA = Sin(ANG)
    Vx = CosA * VEL
    Vy = SinA * VEL

    X = X + Vx
    Y = Y + Vy

    If X < R Then X = R: Vx = -Vx
    If Y < R Then Y = R: Vy = -Vy
    If X > MaxX - R Then X = MaxX - R: Vx = -Vx
    If Y > MaxY - R Then Y = MaxY - R: Vy = -Vy

    ANG = Atan2(Vx, Vy)


    Dx = TX - X
    Dy = TY - Y
    TargetANG = Atan2(Dx, Dy)

    If Dx * Dx + Dy * Dy < (R * R) + 1 Then
        'ì************************************************
        'Target Reached
        ReachTime = CNT

        Select Case CNT \ 7000 Mod 4
            Case 0
                If TX < MaxX \ 2 Then
                    TX = MaxX - 15
                Else
                    TX = 15
                End If
                TY = 100 + (Rnd * (MaxY - 200))
            Case 1
                If MyIDX Mod 2 = 0 Then
                    If TY < MaxY \ 2 Then
                        TY = MaxY - 15
                    Else
                        TY = 15
                    End If
                    TX = 100 + (Rnd * (MaxX - 200))
                Else
                    If TX < MaxX \ 2 Then
                        TX = MaxX - 15
                    Else
                        TX = 15
                    End If
                    TY = 100 + (Rnd * (MaxY - 200))
                End If

            Case 2
                If TX = MaxX \ 2 + Cos(MyIDX / NH * PI2) * MaxY \ 2 Then
                    TX = MaxX \ 2 - Cos(MyIDX / NH * PI2) * MaxY \ 2
                    TY = MaxY \ 2 + Sin(MyIDX / NH * PI2) * MaxY \ 2
                Else
                    TX = MaxX \ 2 + Cos(MyIDX / NH * PI2) * MaxY \ 2
                    TY = MaxY \ 2 - Sin(MyIDX / NH * PI2) * MaxY \ 2
                End If
            Case 3
                If TX < MaxX \ 2 Then
                    TX = MaxX * 2 / 3 + (MyIDX Mod 9) * 16
                Else
                    TX = MaxX * 1 / 3 - (MyIDX Mod 9) * 16
                End If
                TY = MaxY * 1 / 5 + (MyIDX \ 9) * 16
        End Select

skip:
    End If




End Sub
