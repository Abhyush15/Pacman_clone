VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pacman"
   ClientHeight    =   6285
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Gh4Timer 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   5040
      Top             =   3360
   End
   Begin VB.Timer Gh3Timer 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   4440
      Top             =   3360
   End
   Begin VB.PictureBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   3000
      Picture         =   "Pacman.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   10
      Tag             =   "Ghost2"
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   2520
      Picture         =   "Pacman.frx":04F4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   9
      Tag             =   "Ghost2"
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Disappear 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   960
      Top             =   5640
   End
   Begin VB.Frame Message 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   2068
      TabIndex        =   5
      Top             =   2667
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.PictureBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   2040
      Picture         =   "Pacman.frx":09E8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Tag             =   "Ghost2"
      Top             =   120
      Width           =   300
   End
   Begin VB.Timer StartGhost 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   5880
   End
   Begin VB.Timer ChangeMode 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5160
      Top             =   5880
   End
   Begin VB.Timer Gh2Timer 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   5040
      Top             =   2760
   End
   Begin VB.PictureBox Ghost 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000015&
      Height          =   300
      Index           =   1
      Left            =   1560
      Picture         =   "Pacman.frx":0EDC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Tag             =   "Ghost1"
      Top             =   120
      Width           =   300
   End
   Begin VB.Timer Gh1Timer 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   4440
      Top             =   2760
   End
   Begin VB.Timer PacmanTimer 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   1080
      Top             =   2760
   End
   Begin VB.PictureBox Ghost 
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Pacman 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      Picture         =   "Pacman.frx":13D0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox maze 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1095
      Left            =   7920
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MAP(21) As String
Dim Rows As Integer
Dim Cols As Integer
Dim Speed As Integer
Dim Displacement As Integer
Dim Points As Integer
Dim MaxScore As Integer
Dim Mouth As Integer
Dim Grace As Integer
Dim Lives As Integer
Dim Score(21, 21) As Integer
Dim DisappearCount As Integer


'variable for PACMAN
Dim PosLeft As Integer
Dim PosTop As Integer
Dim MaxDistance As Integer
Dim Direction As Integer

Dim Mode As Integer
Dim GhostInt As Integer

Dim Scatter As Integer
Dim Chase As Integer
Dim Start As Integer

Dim modc As Integer

Dim GhLeft As Integer
Dim GhTop As Integer

Dim Gh1Positions(2, 2) As Integer
Dim Gh1Dir As Integer
Dim Gh1MaxDis As Integer

Dim Gh2Positions(2, 2) As Integer
Dim Gh2Dir As Integer
Dim Gh2MaxDis As Integer


'Ghost3 Variables
Dim Gh3Positions(2, 2) As Integer
Dim Gh3Dir As Integer
Dim Gh3MaxDis As Integer


'Ghost4 Variables
Dim Gh4Positions(2, 2) As Integer
Dim Gh4Dir As Integer
Dim Gh4MaxDis As Integer







Public Sub moveGhostRight(object As PictureBox)
    object.Left = object.Left + Displacement
    object.Picture = LoadPicture(App.Path & "\Images\" & object.Tag & "Right.bmp")
End Sub

Public Sub moveGhostLeft(object As PictureBox)
    object.Left = object.Left - Displacement
    object.Picture = LoadPicture(App.Path & "\Images\" & object.Tag & "Left.bmp")
End Sub

Public Sub moveGhostUp(object As PictureBox)
    object.Top = object.Top - Displacement
    object.Picture = LoadPicture(App.Path & "\Images\" & object.Tag & "Up.bmp")
End Sub

Public Sub moveGhostDown(object As PictureBox)
    object.Top = object.Top + Displacement
    object.Picture = LoadPicture(App.Path & "\Images\" & object.Tag & "Down.bmp")
End Sub



Public Sub movePacmanRight()
    Dim pleft As Integer
    Dim ptop As Integer
    Dim k As Integer
    Pacman.Left = Pacman.Left + Displacement
    pleft = Pacman.Left / 300 + 1
    ptop = Pacman.Top / 300 + 1
    k = (ptop - 1) * 21 + pleft
    maze(k).Picture = LoadPicture(App.Path & "\Images\path.bmp")
    Points = Points + Score(ptop, pleft)
    Score(ptop, pleft) = 0
    If (Mouth < 3) Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmor.bmp")
        Mouth = Mouth + 1
    ElseIf Mouth < 6 Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmcr.bmp")
        Mouth = Mouth + 1
    Else
        Mouth = 0
    End If
    If (Points = MaxScore) Then
        Call NextLevel
    End If
End Sub

Public Sub movePacmanLeft()
    Dim pleft As Integer
    Dim ptop As Integer
    
    Pacman.Left = Pacman.Left - Displacement
    pleft = (Pacman.Left / 300) + 1
    ptop = Pacman.Top / 300 + 1
    
    k = (ptop - 1) * 21 + pleft
    maze(k).Picture = LoadPicture(App.Path & "\Images\path.bmp")
    Points = Points + Score(ptop, pleft)
    Score(ptop, pleft) = 0
    If (Mouth < 3) Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmol.bmp")
        Mouth = Mouth + 1
    ElseIf Mouth < 6 Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmcl.bmp")
        Mouth = Mouth + 1
    Else
        Mouth = 0
    End If
    If (Points = MaxScore) Then
        Call NextLevel
    End If
End Sub

Public Sub movePacmanUp()
    Dim pleft As Integer
    Dim ptop As Integer
    Pacman.Top = Pacman.Top - Displacement
    pleft = Pacman.Left / 300 + 1
    ptop = Pacman.Top / 300 + 1
    k = (ptop - 1) * 21 + pleft
    maze(k).Picture = LoadPicture(App.Path & "\Images\path.bmp")
    Points = Points + Score(ptop, pleft)
    Score(ptop, pleft) = 0
    If (Mouth < 3) Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmou.bmp")
        Mouth = Mouth + 1
    ElseIf Mouth < 6 Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmcu.bmp")
        Mouth = Mouth + 1
    Else
        Mouth = 0
    End If
    If (Points = MaxScore) Then
        Call NextLevel
    End If
End Sub

Public Sub movePacmanDown()
    Dim pleft As Integer
    Dim ptop As Integer
    Pacman.Top = Pacman.Top + Displacement
    pleft = Pacman.Left / 300 + 1
    ptop = Pacman.Top / 300 + 1
    k = (ptop - 1) * 21 + pleft
    
    maze(k).Picture = LoadPicture(App.Path & "\Images\path.bmp")
    Points = Points + Score(ptop, pleft)
    Score(ptop, pleft) = 0
    
    If (Mouth < 3) Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmod.bmp")
        Mouth = Mouth + 1
    ElseIf Mouth < 6 Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\pmcd.bmp")
        Mouth = Mouth + 1
    Else
        Mouth = 0
    End If
    If (Points = MaxScore) Then
        Call NextLevel
    End If
End Sub



Function allowed(keycode As Integer) As Integer
    Dim PosLeft As Integer
    Dim PosTop As Integer
    PosLeft = (Pacman.Left / 300) + 1
    PosTop = (Pacman.Top / 300) + 1
    
    allowed = 0
    

    
    Select Case keycode
    
    Case vbKeyRight
        diff = PosTop * 300 - (Pacman.Top + 300)
        If diff > -Grace And diff < Grace Then
            If (PosLeft * 300 = Pacman.Left + 300) Then
                allowed = Mid(MAP(PosTop), PosLeft, 1)
            Else
                allowed = Mid(MAP(PosTop), PosLeft + 1, 1)
            End If
        End If
        If (allowed <> "0") Then
            Pacman.Top = Pacman.Top + diff
        End If
        

        
        
    Case vbKeyLeft
        diff = PosTop * 300 - (Pacman.Top + 300)
        If diff > -Grace And diff < Grace Then
            If (PosLeft * 300 = Pacman.Left + 300) Then
                allowed = Mid(MAP(PosTop), PosLeft - 1, 1)
            Else
                allowed = Mid(MAP(PosTop), PosLeft, 1)
            End If
        End If
        If (allowed <> "0") Then
            Pacman.Top = Pacman.Top + diff
        End If
        
    Case vbKeyUp
        diff = PosLeft * 300 - (Pacman.Left + 300)
        If diff > -Grace And diff < Grace Then
            If (PosTop * 300 = Pacman.Top + 300) Then
                allowed = Mid(MAP(PosTop - 1), PosLeft, 1)
            Else
                allowed = Mid(MAP(PosTop), PosLeft, 1)
            End If
        End If
        If (allowed <> "0") Then
            Pacman.Left = Pacman.Left + diff
        End If
    
    Case vbKeyDown
        diff = PosLeft * 300 - (Pacman.Left + 300)
        If diff > -Grace And diff < Grace Then
            If (PosTop * 300 = Pacman.Top + 300) Then
                allowed = Mid(MAP(PosTop + 1), PosLeft, 1)
            Else
                allowed = Mid(MAP(PosTop), PosLeft, 1)
            End If
        End If
        If (allowed <> "0") Then
            Pacman.Left = Pacman.Left + diff
        End If
        
    End Select
End Function

Function maxdis(keycode As Integer) As Integer
    leftinc = 0
    topinc = 0
    PosLeft = (Pacman.Left / 300) + 1
    PosTop = (Pacman.Top / 300) + 1
    curr = Mid(MAP(PosTop), PosLeft, 1)
    
    
    Select Case keycode
        Case vbKeyLeft
            leftinc = -1
            topinc = 0
            
            
        Case vbKeyRight
            leftinc = 1
            topinc = 0
            
            
        Case vbKeyUp
            leftinc = 0
            topinc = -1
            
        Case vbKeyDown
            leftinc = 0
            topinc = 1
    End Select
    
    Do While (curr <> "0")
        PosLeft = PosLeft + leftinc
        PosTop = PosTop + topinc
        curr = Mid(MAP(PosTop), PosLeft, 1)
    Loop
    If (keycode = vbKeyLeft Or keycode = vbKeyRight) Then
        maxdis = (PosLeft - 1) * 300
    Else
        maxdis = (PosTop - 1) * 300
    End If
End Function



Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
    If (Direction <> keycode) Then
        PacmanTimer.Enabled = False
        Select Case keycode
            Case vbKeyRight
                If (Direction <> vbKeyRight) Then
                    allowd = allowed(keycode)
                    If (allowd <> 0) Then
                        MaxDistance = maxdis(keycode)
                        PacmanTimer.Enabled = True
                    End If
                End If
                
                
            Case vbKeyLeft
                If (Direction <> vbKeyLeft) Then
                    allowd = allowed(keycode)
                    If (allowd <> 0) Then
                        MaxDistance = maxdis(keycode)
                        Direction = vbKeyLeft
                        PacmanTimer.Enabled = True
                    End If
                End If
        
            Case vbKeyUp
                
                If (Direction <> vbKeyUp) Then
                    allowd = allowed(keycode)
                    If (allowd <> 0) Then
                        MaxDistance = maxdis(keycode)
                        Direction = vbKeyUp
                        PacmanTimer.Enabled = True
                    End If
                End If
                
            Case vbKeyDown
                If (Direction <> vbKeyDown) Then
                    allowd = allowed(keycode)
                    If (allowd <> 0) Then
                        MaxDistance = maxdis(keycode)
                        Direction = vbKeyDown
                        PacmanTimer.Enabled = True
                    End If
                End If
    End Select
    Direction = keycode
    End If
End Sub

Private Sub pacmantimer_Timer()
    Select Case Direction
        Case vbKeyRight
            If (Pacman.Left < (MaxDistance - 300)) Then
                Call movePacmanRight
            Else
                PacmanTimer.Enabled = False
            End If
        
        Case vbKeyLeft
            If (Pacman.Left > (MaxDistance + 300)) Then
                Call movePacmanLeft
            Else
                PacmanTimer.Enabled = False
            End If
            
        Case vbKeyUp
            If (Pacman.Top > (MaxDistance + 300)) Then
                Call movePacmanUp
            Else
                PacmanTimer.Enabled = False
            End If
            
        Case vbKeyDown
            If (Pacman.Top < (MaxDistance - 300)) Then
                Call movePacmanDown
            Else
                PacmanTimer.Enabled = False
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i, j As Integer
    Dim k As Integer
    
    Start = 1               'Constants
    Scatter = 2
    Chase = 3
    Speed = 30
    Displacement = 50
    Grace = 155
    Rows = 21
    Cols = 21
    
    
     MAP(1) = "0000000000" + "0" + "0000000000"
     MAP(2) = "0211191190" + "0" + "0911911120"
     MAP(3) = "0100010010" + "0" + "0100100010"
     MAP(4) = "0100010010" + "0" + "0100100010"
     MAP(5) = "0911190091" + "9" + "1900911190"
     MAP(6) = "0100010000" + "1" + "0000100010"
     MAP(7) = "0100010000" + "1" + "0000100010"
     MAP(8) = "0911191111" + "9" + "1111911190"
     
     MAP(9) = "0100010000" + "0" + "0000100010"
    MAP(10) = "0100010055" + "5" + "5500100010"
    MAP(11) = "0100010955" + "9" + "5590100010"
    MAP(12) = "0100010055" + "5" + "5500100010"
    MAP(13) = "0100010000" + "0" + "0000100010"
    
    MAP(14) = "0911191111" + "9" + "1111911190"
    MAP(15) = "0100010000" + "1" + "0000100010"
    MAP(16) = "0100010000" + "1" + "0000100010"
    MAP(17) = "0911190091" + "9" + "1900911190"
    MAP(18) = "0100010010" + "0" + "0100100010"
    MAP(19) = "0100010010" + "0" + "0100100010"
    MAP(20) = "0211191190" + "0" + "0911911120"
    MAP(21) = "0000000000" + "0" + "0000000000"
    
    Level = 1
    
    'Initialize COde
    Call NewGame(True)
    
End Sub

Private Sub NewGame(CreateMap As Boolean)

    Mouth = 0
    Points = 0
    MaxScore = 0
    Lives = 3
    
    Message.Visible = False
    Message.Enabled = False
    Title.Enabled = False
    Text.Enabled = False
    
    k = 1
    For i = 1 To Rows Step 1
        For j = 1 To Cols Step 1
            If CreateMap Then
                Load maze(k)
                maze(k).Left = (j - 1) * 300
                maze(k).Top = (i - 1) * 300
                maze(k).Visible = True
            End If
            Score(i, j) = 0
            Select Case Val(Mid(MAP(i), j, 1))
                
                    Case 0
                        maze(k).Picture = LoadPicture(App.Path & "\Images\Wall.bmp")
                        
                    Case 1
                        maze(k).Picture = LoadPicture(App.Path & "\Images\1.bmp")
                        Score(i, j) = 1
                        MaxScore = MaxScore + 1
                        
                    Case 2
                        maze(k).Picture = LoadPicture(App.Path & "\Images\2.bmp")
                        Score(i, j) = 2
                        MaxScore = MaxScore + 2
                    
                    Case 4
                        maze(k).Picture = LoadPicture(App.Path & "\Images\Wall.bmp")
                    
                    Case 5
                        maze(k).Picture = LoadPicture(App.Path & "\Images\Path.bmp")
                        
                    Case 9
                        If (i = 11) Then
                            maze(k).Picture = LoadPicture(App.Path & "\Images\Path.bmp")
                        Else
                            maze(k).Picture = LoadPicture(App.Path & "\Images\1.bmp")
                            Score(i, j) = 1
                            MaxScore = MaxScore + 1
                        End If
                End Select
                k = k + 1
        Next j
    Next i
    
    Call Reset
    
End Sub


Private Sub Reset()
    
    ChangeMode.Enabled = False
    Disappear.Enabled = False
    DisappearCount = 15
    
    Pacman.Picture = LoadPicture(App.Path & "\Images\pmcr.bmp")
    
    Direction = 36
    Pacman.Left = 300 * 10
    Pacman.Top = 300 * 13
    
    
    modc = 2
    GhostInt = 0
    Mode = Scatter
    
    Gh1Dir = vbKeyDown
    Gh1MaxDis = 0
    Ghost(1).Left = 300 * 7
    Ghost(1).Top = 300 * 10
    Gh1Positions(Scatter, 0) = 10
    Gh1Positions(Scatter, 1) = 7
    
    Gh2Dir = vbKeyDown
    Gh2MaxDis = 0
    Ghost(2).Left = 300 * 13
    Ghost(2).Top = 300 * 10
    Gh2Positions(Scatter, 0) = 10
    Gh2Positions(Scatter, 1) = 7
    
    If (Level = 2) Then
        'Level 2 Code Comes here
    End If
    
    PacmanTimer.Interval = Speed
    Gh1Timer.Interval = Speed
    Gh2Timer.Interval = Speed
    
    PacmanTimer.Enabled = True
    StartGhost.Enabled = True
    
End Sub


Function nextmove(object As PictureBox, ByRef ghostdir As Integer) As Integer
    
    Dim topinc As Integer
    Dim leftinc As Integer
    
    Dim topdiff As Integer
    Dim leftdiff As Integer
    
    If Mode = Chase Then
        PosLeft = (Pacman.Left / 300) + 1
        PosTop = (Pacman.Top / 300) + 1
    ElseIf Mode = Scatter And object.Index = 1 Then
        PosLeft = Gh1Positions(Scatter, 0)
        PosTop = Gh1Positions(Scatter, 1)
    ElseIf Mode = Scatter And object.Index = 2 Then
        PosLeft = Gh2Positions(Scatter, 0)
        PosTop = Gh2Positions(Scatter, 1)
    End If
    
    GhLeft = (object.Left / 300) + 1
    GhTop = (object.Top / 300) + 1
    
    
    allowdRight = Mid(MAP(GhTop), GhLeft + 1, 1)
    allowdleft = Mid(MAP(GhTop), GhLeft - 1, 1)
    
    allowdDown = Mid(MAP(GhTop + 1), GhLeft, 1)
    allowdup = Mid(MAP(GhTop - 1), GhLeft, 1)
    
    topdiff = Abs(PosTop - GhTop)
    leftdiff = Abs(PosLeft - GhLeft)
    
    If (topdiff >= leftdiff) Then
        Select Case ghostdir
        
                Case vbKeyRight
                    If (allowdup <> "0" And PosTop <= GhTop) Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf (allowdDown <> "0" And PosTop > GhTop) Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdRight <> "0" Then
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    End If
                    
                Case vbKeyLeft
                    If (allowdup <> "0" And PosTop <= GhTop) Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf (allowdDown <> "0" And PosTop > GhTop) Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdleft <> "0" Then
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    End If
                    
                    
                Case vbKeyUp
                    If (allowdup <> "0" And PosTop <= GhTop) Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf allowdleft <> "0" And PosLeft <= GhLeft Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" And PosLeft > GhLeft Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdleft <> "0" Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    End If
                    
                    
                    
                Case vbKeyDown
                    If (allowdDown <> "0" And PosTop >= GhTop) Then
                        ghostdir = vbKeyDown
                        topinc = 1
                        leftinc = 0
                    ElseIf allowdleft <> "0" And PosLeft <= GhLeft Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" And PosLeft > GhLeft Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdleft <> "0" Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        topinc = 1
                        leftinc = 0
                    End If
        End Select
    Else
        Select Case ghostdir
        
                Case vbKeyUp
                    If (allowdleft <> "0" And PosLeft <= GhLeft) Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" And PosLeft > GhLeft Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        leftinc = 0
                        topinc = -1
                    ElseIf allowdleft <> "0" Then
                        ghostdir = vbKeyLeft
                        topinc = 0
                        leftinc = -1
                    ElseIf allowdRight <> "0" Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    End If
                    
                Case vbKeyDown
                    If (allowdleft <> "0" And PosLeft <= GhLeft) Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf allowdRight <> "0" And PosLeft > GhLeft Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdleft <> "0" Then
                        ghostdir = vbKeyLeft
                        topinc = 0
                        leftinc = -1
                    ElseIf allowdRight <> "0" Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    End If
                    
                    
                Case vbKeyLeft
                    If (allowdleft <> "0" And PosLeft <= GhLeft) Then
                        ghostdir = vbKeyLeft
                        leftinc = -1
                        topinc = 0
                    ElseIf (allowdup <> "0" And PosTop <= GhTop) Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf (allowdDown <> "0" And PosTop > GhTop) Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdleft <> "0" Then
                        leftinc = -1
                        topinc = 0
                    End If
                    
                Case vbKeyRight
                    If (allowdRight <> "0" And PosLeft >= GhLeft) Then
                        ghostdir = vbKeyRight
                        leftinc = 1
                        topinc = 0
                    ElseIf (allowdup <> "0" And PosTop <= GhTop) Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf (allowdDown <> "0" And PosTop > GhTop) Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdup <> "0" Then
                        ghostdir = vbKeyUp
                        topinc = -1
                        leftinc = 0
                    ElseIf allowdDown <> "0" Then
                        ghostdir = vbKeyDown
                        leftinc = 0
                        topinc = 1
                    ElseIf allowdRight <> "0" Then
                        leftinc = -1
                        topinc = 0
                    End If
        End Select
    End If
    GhTop = GhTop + topinc
    GhLeft = GhLeft + leftinc
    While ((Mid(MAP(GhTop), GhLeft, 1) <> "9") And (Mid(MAP(GhTop), GhLeft, 1) <> "2"))
        GhTop = GhTop + topinc
        GhLeft = GhLeft + leftinc
    Wend
    nextmove = (GhTop * Abs(topinc) + GhLeft * Abs(leftinc) - 1) * 300
    
   
End Function

Private Sub Gh1Timer_Timer()
    Select Case Gh1Dir
        Case vbKeyRight
            If (Ghost(1).Left < (Gh1MaxDis)) Then
                Call moveGhostRight(Ghost(1))
            Else
                'Gh1Timer.Enabled = False
                Gh1MaxDis = nextmove(Ghost(1), Gh1Dir)
                'Gh1Timer.Enabled = True
            End If
        
        Case vbKeyLeft
            If (Ghost(1).Left > (Gh1MaxDis)) Then
                Call moveGhostLeft(Ghost(1))
            Else
                'Gh1Timer.Enabled = False
                Gh1MaxDis = nextmove(Ghost(1), Gh1Dir)
                'Gh1Timer.Enabled = True
            End If
            
        Case vbKeyUp
            If (Ghost(1).Top > (Gh1MaxDis)) Then
                Call moveGhostUp(Ghost(1))
            Else
                'Gh1Timer.Enabled = False
                Gh1MaxDis = nextmove(Ghost(1), Gh1Dir)
                'Gh1Timer.Enabled = True
            End If
            
        Case vbKeyDown
            If (Ghost(1).Top < (Gh1MaxDis)) Then
                Call moveGhostDown(Ghost(1))
            Else
                'Gh1Timer.Enabled = False
                Gh1MaxDis = nextmove(Ghost(1), Gh1Dir)
                'Gh1Timer.Enabled = True
            End If
    End Select
    If Abs(Pacman.Left - Ghost(1).Left) < 200 And Abs(Pacman.Top - Ghost(1).Top) < 200 Then
        Call LifeLost
    End If
End Sub


Private Sub Gh2Timer_Timer()
    Select Case Gh2Dir
        Case vbKeyRight
            If (Ghost(2).Left < (Gh2MaxDis)) Then
                Call moveGhostRight(Ghost(2))
            Else
                'Gh2Timer.Enabled = False
                Gh2MaxDis = nextmove(Ghost(2), Gh2Dir)
                'Gh2Timer.Enabled = True
            End If
        
        Case vbKeyLeft
            If (Ghost(2).Left > (Gh2MaxDis)) Then
                Call moveGhostLeft(Ghost(2))
            Else
                'Gh2Timer.Enabled = False
                Gh2MaxDis = nextmove(Ghost(2), Gh2Dir)
                'Gh2Timer.Enabled = True
            End If
            
        Case vbKeyUp
            If (Ghost(2).Top > (Gh2MaxDis)) Then
                Call moveGhostUp(Ghost(2))
            Else
                'Gh2Timer.Enabled = False
                Gh2MaxDis = nextmove(Ghost(2), Gh2Dir)
                'Gh2Timer.Enabled = True
            End If
            
        Case vbKeyDown
            If (Ghost(2).Top < (Gh2MaxDis)) Then
                Call moveGhostDown(Ghost(2))
            Else
                'Gh2Timer.Enabled = False
                Gh2MaxDis = nextmove(Ghost(2), Gh2Dir)
                'Gh2Timer.Enabled = True
            End If
    End Select
    If Abs(Pacman.Left - Ghost(2).Left) < 200 And Abs(Pacman.Top - Ghost(2).Top) < 200 Then
        Call LifeLost
    End If
End Sub



Private Sub NextLevel()

    PacmanTimer.Enabled = False
    PacmanTimer.Interval = 0
    ChangeMode.Enabled = False
    Gh1Timer.Enabled = False
    Gh2Timer.Enabled = False
    Level = Level + 1
    Title.Caption = "You Won!"
    Text.Caption = "Click Here..."
    Title.Enabled = True
    Text.Enabled = True
    Message.Enabled = True
    Message.Visible = True
End Sub


Private Sub LifeLost()
    PacmanTimer.Enabled = False
    PacmanTimer.Interval = 0
    ChangeMode.Enabled = False
    Gh1Timer.Enabled = False
    Gh2Timer.Enabled = False
    Lives = Lives - 1
    If Lives > 0 Then
        Disappear.Enabled = True
    Else
        Disappear.Enabled = True
        Title.Caption = "You Lost!"
        Text.Caption = "Click Here..."
        Title.Enabled = True
        Text.Enabled = True
        Message.Enabled = True
        Message.Visible = True
    End If
End Sub

Private Sub Disappear_Timer()
    
    If (DisappearCount >= 0 And DisappearCount < 10) Then
        Pacman.Picture = LoadPicture(App.Path & "\Images\Disappear" & DisappearCount & ".jpg")
        DisappearCount = DisappearCount - 1
    ElseIf DisappearCount > -30 Then
        DisappearCount = DisappearCount - 1
    ElseIf Lives > 0 Then
        Call Reset
    End If
End Sub


Private Sub ChangeMode_Timer()
    StartGhost.Enabled = False
    If modc < 3 Then
        modc = modc + 1
        Mode = Scatter
    ElseIf modc > 7 Then
        modc = 0
    Else
        modc = modc + 1
        Mode = Chase
    End If
End Sub
Private Sub StartGhost_Timer()
    If GhostInt = 2 Then
        Gh1Timer.Enabled = True
    ElseIf GhostInt = 3 Then
        Gh1Positions(Scatter, 0) = 1
        Gh1Positions(Scatter, 1) = 1
        Gh2Timer.Enabled = True
    ElseIf GhostInt = 4 Then
        Gh2Positions(Scatter, 0) = 19
        Gh2Positions(Scatter, 1) = 19
    ElseIf GhostInt > 4 Then
        ChangeMode.Enabled = True
    End If
    GhostInt = GhostInt + 1
End Sub

Private Sub Text_Click()
    Call NewGame(False)
End Sub

Private Sub Title_Click()
    Call NewGame(False)
End Sub
