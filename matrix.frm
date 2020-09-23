VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Matrix By Kevin Pfister"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   10110
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Matrix"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "matrix.frx":0000
   ScaleHeight     =   40.25
   ScaleMode       =   4  'Character
   ScaleWidth      =   84.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrMain 
      Interval        =   20
      Left            =   90
      Top             =   90
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LengthOfDrop(1 To 200, 1 To 200) As Byte    'Length of Drop
Dim Leading(1 To 200, 1 To 200) As Byte    'Is it a leading one
Dim Letter(1 To 200, 1 To 200) As Byte    'Letter
Dim Colour(1 To 200, 1 To 200) As Integer    'Colour of the letter /symbol
Dim WaitBeforeClear(1 To 200, 1 To 200) As Byte        'Wait before it dissappears
Dim MaxLength   'Max Length
Dim MaxWait   'Max Wait
Dim H   'No of Drops
Dim M   'Fade Speed
Dim O   'Fall From Top
Dim Q   'Fade
Dim S
Dim U
Dim X1
Dim Y1
Dim LastXPos
Dim LastYPos

Private Sub Form_Click()
    ShowCursor (1)
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ShowCursor (1)
    End
End Sub

Private Sub Form_Load()
    ShowCursor (0)
    FrmMain.ForeColor = RGB(0, 220, 0)
    MaxLength = GetSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", 100)
    MaxWait = GetSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", 30)
    H = GetSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", 20)
    M = GetSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", 4)
    O = GetSetting("Kevin Pfister's Matrix", "Options", "FromTop", 1)
    Q = GetSetting("Kevin Pfister's Matrix", "Colour", "Fade", 0)
    S = GetSetting("Kevin Pfister's Matrix", "Colour", "MColours", 1)
    U = GetSetting("Kevin Pfister's Matrix", "Colour", "HighLights", 1)
    XY = GetSetting("Kevin Pfister's Matrix", "Res", "XY", "1024 by 768 pixels")
    If XY = "640 by 480 pixels" Then
        X1 = 640
        Y1 = 480
    ElseIf XY = "720 by 480 pixels" Then
        X1 = 720
        Y1 = 480
    ElseIf XY = "720 by 576 pixels" Then
        X1 = 720
        Y1 = 576
    ElseIf XY = "800 by 600 pixels" Then
        X1 = 800
        Y1 = 600
    ElseIf XY = "1024 by 768 pixels" Then
        X1 = 1024
        Y1 = 768
    ElseIf XY = "1152 by 864 pixels" Then
        X1 = 1152
        Y1 = 864
    ElseIf XY = "1280 by 960 pixels" Then
        X1 = 1280
        Y1 = 960
    ElseIf XY = "1280 by 1024 pixels" Then
        X1 = 1280
        Y1 = 1024
    ElseIf XY = "1600 by 1200 pixels" Then
        X1 = 1600
        Y1 = 1200
    End If
    Randomize Timer
    For DoRand = 1 To H
        XR = Int(Rnd * (0.12109375 * X1)) + 1
        YR = Int(Rnd * (0.06770833 * Y1)) + 1
        LengthOfDrop(XR, YR) = Int(Rnd * MaxLength)
        Leading(XR, YR) = 1
        Letter(XR, YR) = Int(Rnd * 43) + 65
        If S = 1 Then
            Colour(XR, YR) = Int(Rnd * 200) + 55
        End If
    Next
End Sub

Sub OneColour()
    FrmMain.Cls
    For X = 1 To (0.12109375 * X1)
        For Y = 1 To (0.06770833 * Y1)
            If Leading(X, Y) = 1 Then 'Is it leading
                If Y + 1 <= 200 Then 'Is it smaller than the screen height
                    If LengthOfDrop(X, Y) > 0 Then 'Is there still drops in this column
                        LengthOfDrop(X, Y + 1) = LengthOfDrop(X, Y) - 1
                        Leading(X, Y + 1) = 2
                        Letter(X, Y + 1) = Int(Rnd * 43) + 65
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = MaxWait
                    Else    'End of Drop(Kill Letter/Symbol)
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = MaxWait
                    End If
                Else    'End of Drop(Kill Letter/Symbol)
                    Leading(X, Y) = 0
                    WaitBeforeClear(X, Y) = MaxWait
                End If
            ElseIf WaitBeforeClear(X, Y) > 0 Then 'Is the Letter/Symbol dieing?
                WaitBeforeClear(X, Y) = WaitBeforeClear(X, Y) - 1
                If WaitBeforeClear(X, Y) = 0 Then
                    Letter(X, Y) = 0
                End If
            End If
            If Leading(X, Y) = 1 Or Leading(X, Y) = 2 Then
                Leading(X, Y) = 1
                Drops = Drops + 1
            End If
            If Letter(X, Y) > 0 Then
                FrmMain.CurrentX = X
                FrmMain.CurrentY = Y - 5
                If Leading(X, Y) = 0 Or U = 0 Then
                    FrmMain.Print Chr(Letter(X, Y))
                Else
                    FrmMain.ForeColor = vbWhite
                    FrmMain.Print Chr(Letter(X, Y))
                    FrmMain.ForeColor = RGB(0, 220, 0)
                End If
            End If
        Next
    Next
    If Drops < H Then
        For MakeNew = Drops To H
            XR = Int(Rnd * (0.12109375 * X1)) + 1
            If O = 1 Then
                YR = Int(Rnd * 5) + 1
            Else
                YR = Int(Rnd * (0.06770833 * Y1)) + 1
            End If
            LengthOfDrop(XR, YR) = Int(Rnd * MaxLength)
            Leading(XR, YR) = 1
            Letter(XR, YR) = 64 + Int(Rnd * 26)
        Next
    End If
End Sub

Sub MoreThanOneColour()
    FrmMain.Cls
    For X = 1 To (0.12109375 * X1)
        For Y = 1 To (0.06770833 * Y1)
            If Leading(X, Y) = 1 Then 'Is it leading
                If Y + 1 <= 200 Then 'Is it smaller than the screen height
                    If LengthOfDrop(X, Y) > 0 Then 'Is there still drops in this column
                        LengthOfDrop(X, Y + 1) = LengthOfDrop(X, Y) - 1
                        Leading(X, Y + 1) = 2
                        Letter(X, Y + 1) = Int(Rnd * 43) + 65
                        Colour(X, Y + 1) = Int(Rnd * 200) + 55
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = MaxWait
                    Else    'End of Drop(Kill Letter/Symbol)
                        Leading(X, Y) = 0
                        WaitBeforeClear(X, Y) = MaxWait
                    End If
                Else    'End of Drop(Kill Letter/Symbol)
                    Leading(X, Y) = 0
                    WaitBeforeClear(X, Y) = MaxWait
                End If
            ElseIf WaitBeforeClear(X, Y) > 0 Then 'Is the Letter/Symbol dieing?
                WaitBeforeClear(X, Y) = WaitBeforeClear(X, Y) - 1
                If Q = 1 Then   'Is fading ativated
                    Colour(X, Y) = Colour(X, Y) - M
                End If
                If WaitBeforeClear(X, Y) = 0 Or Colour(X, Y) < 0 Then
                    Letter(X, Y) = 0
                End If
            End If
            If Leading(X, Y) = 1 Or Leading(X, Y) = 2 Then
                Leading(X, Y) = 1
                Drops = Drops + 1
            End If
            If Letter(X, Y) > 0 Then
                FrmMain.CurrentX = X
                FrmMain.CurrentY = Y - 5
                If (Leading(X, Y) = 0 Or U = 0) And S = 1 Then
                    FrmMain.ForeColor = RGB(0, Colour(X, Y), 0)
                    FrmMain.Print Chr(Letter(X, Y))
                Else
                    FrmMain.ForeColor = vbWhite
                    FrmMain.Print Chr(Letter(X, Y))
                End If
            End If
        Next
    Next
    If Drops < H Then
        For MakeNew = Drops To H
            XR = Int(Rnd * (0.12109375 * X1)) + 1
            If O = 1 Then
                YR = Int(Rnd * 5) + 1
            Else
                YR = Int(Rnd * (0.06770833 * Y1)) + 1
            End If
            LengthOfDrop(XR, YR) = Int(Rnd * MaxLength)
            Leading(XR, YR) = 1
            Letter(XR, YR) = 64 + Int(Rnd * 26)
            Colour(XR, YR) = Int(Rnd * 200) + 55
        Next
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (LastXPos = 0 And LastY = 0) Or (Abs(LastXPos - X) < 2 And Abs(LastYPos - Y) < 2) Then
        LastXPos = X
        LastYPos = Y
    Else
        ShowCursor (1)
        End
    End If
End Sub

Private Sub TmrMain_Timer()
    If S = 0 Then
        Call OneColour
    Else
        Call MoreThanOneColour
    End If
End Sub
