VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "Reciprocal Velocity Obstacle"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chSaveJpg 
      Caption         =   "Save Frame"
      Height          =   855
      Left            =   14520
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox chAuto 
      Caption         =   "Auto Pan"
      Height          =   375
      Left            =   14520
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtNH 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   14520
      TabIndex        =   11
      Text            =   "100"
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox chFollow 
      Caption         =   "Follow"
      Height          =   375
      Left            =   14520
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "debug"
      Height          =   375
      Left            =   14520
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   60
      ScaleHeight     =   705
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   921
      TabIndex        =   1
      Top             =   120
      Width           =   13815
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   12840
         TabIndex        =   9
         Top             =   9840
         Width           =   255
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   13080
         TabIndex        =   8
         Top             =   9720
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   12360
         TabIndex        =   7
         Top             =   9720
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   11760
         TabIndex        =   6
         Top             =   9720
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   10560
         TabIndex        =   5
         Top             =   9720
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   11160
         TabIndex        =   4
         Top             =   10080
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   11160
         TabIndex        =   3
         Top             =   9480
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(re)Start"
      Height          =   735
      Left            =   14520
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   14520
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
'***********************************************************************************

Option Explicit

Private Sub chAuto_Click()
    If chAuto Then chFollow.Value = vbUnchecked

End Sub

Private Sub chFollow_Click()
    If chFollow Then chAuto.Value = vbUnchecked
End Sub

Private Sub cmdNAVIGATE_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = True

    Do

        Select Case Index
            Case 0
                PanZoomChanged = True
                PanY = PanY - 0.001 / Zoom

            Case 1
                PanZoomChanged = True
                PanY = PanY + 0.001 / Zoom


            Case 2
                PanZoomChanged = True
                PanX = PanX - 0.001 / Zoom

            Case 3
                PanZoomChanged = True
                PanX = PanX + 0.001 / Zoom


            Case 4
                PanZoomChanged = True
                Zoom = Zoom / 1.000002
            Case 5
                PanZoomChanged = True
                Zoom = Zoom * 1.000002    '1.2

            Case 6
                PanZoomChanged = True
                Zoom = 1
                PanX = MaxX / 2
                PanY = MaxY / 2
                Zoom = PIC.Height / MaxY
        End Select

        '        SW.DRAW
        '        PIC.Refresh

        DoEvents
        InvZoom = 1 / Zoom
    Loop While Navigating

End Sub

Private Sub cmdNAVIGATE_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = False

End Sub

Private Sub Command1_Click()
    Dim I          As Long


    Randomize Timer

    'AddLine 30, 30, 100, 35
    'AddLine 600, 300, 622, 250
    'AddLine MaxX \ 2, MaxY \ 2 - 100, MaxX \ 2, MaxY \ 2 + 100

    NH = Val(txtNH)

    ReDim H(NH)

    For I = 1 To NH
        With H(I)
            .InitTrail 50
            .X = Rnd * MaxX
            .Y = Rnd * MaxY
            .X = MaxX \ 2 - Cos(I / NH * PI2) * MaxY \ 2
            .Y = MaxY \ 2 + Sin(I / NH * PI2) * MaxY \ 2

            .TX = MaxX \ 2 + Cos(I / NH * PI2) * MaxY \ 2
            .TY = MaxY \ 2 - Sin(I / NH * PI2) * MaxY \ 2

            .ANG = Rnd * PI2
            .ANG = -I / NH * PI2

            .DesiredANG = .ANG
            .R = 2 + Rnd * 1
            .Shyness = Rnd
            .MyColor = RGB(255 - .Shyness * 255, .Shyness * 255, 150)
            .MaxVEL = (1.33 * 0.5) * (0.9 + Rnd * 0.2)

            .DesiredVEL = .MaxVEL
            If I = 2 Then
                .MaxVEL = 0.1: .DesiredVEL = 0: .R = 10:    '.X = MaxX \ 2: .Y = MaxY \ 2
            End If
        End With
    Next

    Do
        allRVO
        If CNT Mod RenderFrame = 0 Then BitBlt PIC.hdc, 0, 0, MaxX \ 1, MaxY \ 1, PIC.hdc, 0, 0, vbBlack

        For I = 1 To NH
            H(I).MOVE I
        Next
        If CNT Mod RenderFrame = 0 Then
            DrawGrid PIC.hdc
            DrawLines PIC.hdc
            For I = 1 To NH
                H(I).RenderTrail PIC.hdc
            Next
            For I = 1 To NH
                H(I).Render PIC.hdc
            Next
            If chFollow Then
                PanX = PanX * 0.9 + H(1).X * 0.1
                PanY = PanY * 0.9 + H(1).Y * 0.1
            End If
            If chAuto Then CalcAuto

            If chSaveJpg Then
                SaveJPG PIC.Image, App.Path & "\Frames\" & Format(FR, "00000") & ".jpg", 90
                FR = FR + 1
                Label1.Caption = FR & "   " & Int(FR / 30)
            End If

            PIC.Refresh
        End If
        DoEvents
        CNT = CNT + 1
    Loop While True



End Sub
Private Sub CalcAuto()
    Dim I          As Long
    Dim X          As Single
    Dim Y          As Single
    Dim x1         As Long
    Dim y1         As Long

    Dim SomeOutSide As Boolean

    For I = 1 To NH
        X = X + H(I).X
        Y = Y + H(I).Y
    Next
    X = X / NH
    Y = Y / NH
    PanX = X
    PanY = Y
    'Zoom = 10
    'Do
    '    SomeOutSide = False
    '    Zoom = Zoom * 0.9
    '    For I = 1 To NH
    '        X1 = XtoScreen(H(I).X)
    '        Y1 = YtoScreen(H(I).Y)
    '        If Not (IsInsideScreen(X1, Y1)) Then SomeOutSide = True: Exit For
    '    Next
    'Loop While SomeOutSide


End Sub


Private Sub Form_Load()

    PIC.Height = 720
    PIC.Width = Int(4 / 3 * PIC.Height)

    If Dir(App.Path & "\Frames", vbDirectory) = "" Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.jpg") <> "" Then Kill App.Path & "\Frames\*.jpg"


    MaxXPic = PIC.Width
    MaxYPic = PIC.Height
    CenX = MaxXPic \ 2
    CenY = MaxYPic \ 2

    MaxX = PIC.Width * 1
    MaxY = PIC.Height * 1

    PanX = MaxX \ 2               'CenX
    PanY = MaxY \ 2               'CenY
    Zoom = 1
    InvZoom = 1 / Zoom



    '-------------------------------------------
    PanX = MaxX / 2
    PanY = MaxY / 2
    Zoom = PIC.Height / MaxY

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub




