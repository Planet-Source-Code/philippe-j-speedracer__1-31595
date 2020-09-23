VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeedRacer"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNext 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   3120
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.Timer TimeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   960
   End
   Begin VB.TextBox txtCars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Text            =   "2:00"
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   7560
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   7560
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox picCar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   5040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   4920
      Width           =   465
   End
   Begin VB.Timer lightTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   3360
   End
   Begin VB.PictureBox PicRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2160
      ScaleHeight     =   405
      ScaleWidth      =   570
      TabIndex        =   2
      Top             =   -120
      Width           =   570
   End
   Begin VB.PictureBox PicLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1320
      ScaleHeight     =   405
      ScaleWidth      =   570
      TabIndex        =   1
      Top             =   -120
      Width           =   570
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cars to Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Cleared Cars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Time left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgRightCar 
      Height          =   960
      Index           =   3
      Left            =   6360
      Picture         =   "Form1.frx":04CF
      Top             =   2880
      Width           =   465
   End
   Begin VB.Image imgLeftCar 
      Height          =   960
      Index           =   3
      Left            =   6840
      Picture         =   "Form1.frx":099E
      Top             =   2880
      Width           =   465
   End
   Begin VB.Image imgLeftCar 
      Height          =   960
      Index           =   2
      Left            =   6840
      Picture         =   "Form1.frx":0E6D
      Top             =   1920
      Width           =   465
   End
   Begin VB.Image imgRightCar 
      Height          =   960
      Index           =   2
      Left            =   6360
      Picture         =   "Form1.frx":135D
      Top             =   1920
      Width           =   465
   End
   Begin VB.Image imgLeftCar 
      Height          =   960
      Index           =   1
      Left            =   6840
      Picture         =   "Form1.frx":1851
      Top             =   960
      Width           =   465
   End
   Begin VB.Image imgRightCar 
      Height          =   960
      Index           =   1
      Left            =   6360
      Picture         =   "Form1.frx":1D83
      Top             =   960
      Width           =   465
   End
   Begin VB.Image imgLeftCar 
      Height          =   960
      Index           =   0
      Left            =   6840
      Picture         =   "Form1.frx":22B4
      Top             =   0
      Width           =   465
   End
   Begin VB.Image imgRightCar 
      Height          =   960
      Index           =   0
      Left            =   6360
      Picture         =   "Form1.frx":27C5
      Top             =   0
      Width           =   465
   End
   Begin VB.Shape light3 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Light2 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   255
   End
   Begin VB.Shape Light1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape LightFrame 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3360
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   7095
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   7095
      Left            =   2760
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image imgCar 
      Height          =   960
      Left            =   6840
      Picture         =   "Form1.frx":2CD4
      Top             =   3960
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   6240
      Picture         =   "Form1.frx":31A3
      Top             =   4200
      Width           =   570
   End
   Begin VB.Image imgLine5 
      Height          =   3225
      Left            =   6120
      Picture         =   "Form1.frx":45D5
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgLine4 
      Height          =   3225
      Left            =   5880
      Picture         =   "Form1.frx":5D9B
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgLine3 
      Height          =   3225
      Left            =   5640
      Picture         =   "Form1.frx":7561
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgLine2 
      Height          =   3225
      Left            =   5400
      Picture         =   "Form1.frx":8D27
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgLine1 
      Height          =   3225
      Left            =   5160
      Picture         =   "Form1.frx":A4ED
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgRoadLine 
      Height          =   7065
      Left            =   1920
      Picture         =   "Form1.frx":BCB3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuThehelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Year: 2002
'Author: Philippe Johnston
'Alias: Evilstrike
'Alias: Grunty
'program: SpeedRacer
'email: evilstrike@videotron.ca
'purpose: Pure enjoyment
'Motive: I dream it, I code it

'The pictures were taken from http://trophy.sourceforge.net/index.php3?body=screenshots
'The have not been drawn by we and do not belong to me
'I will not take responsability of what you do with the pictures
'Sorry of the game turns out to be addictive because that was
'the main purpose :o)
'Enjoy

Option Explicit
Dim gameover As Boolean
Dim SelectSide As Integer

Dim RoadSpeed As Single
Dim LeftObjectSpeed As Single
Dim RightObjectSpeed As Single
Dim CarSpeed As Single

Dim lightcolor As Integer
Dim TheTag As String
Dim CurrentTime As Integer
Dim minute As Integer
Dim second As Integer
Dim TimeTaken As Integer

Dim Win As Integer
Dim NextWin As Integer
Dim LeftCar As Boolean
Dim RightCar As Boolean

Private Sub Form_Activate()
    'set the coordinates of everything
    xcord = 144
    ycord = 400
    ObLxcord = 88
    ObLycord = 0
    ObRxcord = 144
    ObRycord = 0
    'set the speed of the road (loop speed)
    RoadSpeed = 50
    'set the speed of the right lane car
    RightObjectSpeed = 10
    CarSpeed = 5
    'set the next win
    NextWin = 5
End Sub

Private Sub Command1_Click()
    'set the coordinates of everything
    xcord = 144
    ycord = 400
    ObLxcord = 88
    ObLycord = 0
    ObRxcord = 144
    ObRycord = 0
    'set road speed
    RoadSpeed = 50
    'set speed of right lane car
    RightObjectSpeed = 10
    'Set the light timer to start at red light
    lightcolor = 0
    'clear number of cars
    txtCars.Text = ""
    'set minute time
    minute = 1
    'set second time
    second = 59
    'set timer text
    txtTime.Text = "2:00"
    'set the amout of cars to pass to win
    Win = NextWin
    txtNext.Text = Win
    Command1.Enabled = False
    gameover = False
    Light1.Visible = True
    Light2.Visible = True
    light3.Visible = True
    LightFrame.Visible = True
    RightCar = False
    LeftCar = False
    lightTimer.Enabled = True
    'set lights to black
    Light1.FillColor = &H0&
    Light2.FillColor = &H0&
    light3.FillColor = &H0&
    'clear all picturebox
    PicLeft.Picture = LoadPicture()
    PicRight.Picture = LoadPicture()
    picCar.Picture = LoadPicture()
End Sub

Private Sub lightTimer_Timer()

'timer which will pass from red light to yellow to green
'and then game loop will start
Select Case lightcolor
    Case 0
        Light1.FillColor = &HFF&
        Light2.FillColor = &H0&
        light3.FillColor = &H0&
    Case 1
        Light1.FillColor = &H0&
        Light2.FillColor = &HFFFF&
        light3.FillColor = &H0&
    Case 2
        Light1.FillColor = &H0&
        Light2.FillColor = &H0&
        light3.FillColor = &HFF00&
    Case 3
        Light1.FillColor = &H0&
        Light2.FillColor = &H0&
        light3.FillColor = &H0&
        Light1.Visible = False
        Light2.Visible = False
        light3.Visible = False
        LightFrame.Visible = False
        lightTimer.Enabled = False
        TimeTimer.Enabled = True
        'Start the loop
        Call MasterGameLoop
End Select
    lightcolor = lightcolor + 1

End Sub


Public Sub MasterGameLoop()
Dim nTime As Long
Dim currImage As Integer
Dim inBackOf As Boolean
Do
    'set time before next loop start (depending on the speed)
    nTime = timeGetTime + RoadSpeed
    picCar.Top = ycord
    picCar.Left = xcord
    picCar.Picture = imgCar.Picture
    'right key
    If GetAsyncKeyState(39) < 0 Then
        If xcord < 144 Then
            xcord = xcord + 5
        End If
    End If
    'left key
    If GetAsyncKeyState(37) < 0 Then
        If xcord > 88 Then
            xcord = xcord - 5
        End If
    End If
    'down key
    If GetAsyncKeyState(40) < 0 Then
        RoadSpeed = RoadSpeed + 5
        RightObjectSpeed = RightObjectSpeed - 1
        LeftObjectSpeed = LeftObjectSpeed + 1
    End If
    'up key
    If GetAsyncKeyState(38) < 0 Then
        If RoadSpeed > 2 Then
            RoadSpeed = RoadSpeed - 4
            If RightObjectSpeed < 50 Then
                RightObjectSpeed = RightObjectSpeed + 1
            End If
        End If
    End If
    
    'Give road line to allusion of moving
    Select Case currImage
        Case 1
            imgRoadLine.Picture = imgLine5.Picture
        Case 2
            imgRoadLine.Picture = imgLine4.Picture
        Case 3
            imgRoadLine.Picture = imgLine3.Picture
        Case 4
            imgRoadLine.Picture = imgLine2.Picture
        Case 5
            imgRoadLine.Picture = imgLine1.Picture
            currImage = 0
    End Select
        currImage = currImage + 1
            
    'Place cars
    Call PlaceCars
    
    'if there is already a car in the left lane do the following
    If RightCar = True Then
        'check for collision with right lane car
        If Within(ObRxcord, ObRycord, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
        If Within(ObRxcord + PicRight.Width, ObRycord, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
        If Within(ObRxcord, ObRycord + PicRight.Height, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
        If Within(ObRxcord + PicRight.Width, ObRycord + PicRight.Height, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
        
        ObRycord = ObRycord + RightObjectSpeed
        PicRight.Top = ObRycord
        'check for car offscreen
        'if car is offscreen then update pass cars
            If PicRight.Top >= 500 Then
                RightCar = False
                ObRycord = 0
                RightObjectSpeed = 10
                txtCars.Text = Val(txtCars.Text) + 1
                    'if the correct amount of cars were passed
                    'then end game as a winner
                    If Val(txtCars.Text) = Win Then
                        Dim Winner As Boolean
                            Winner = True
                            gameover = True
                    End If
                'call function to place cars
                Call PlaceCars
            End If
    End If
    
    'if there is already a car in the left lane then do the following
    If LeftCar = True Then
        'if car is not offscreen then it will move downwards depending on speed
        If PicLeft.Top <= 500 Then
            ObLycord = ObLycord + LeftObjectSpeed
            PicLeft.Top = ObLycord
        Else
            'if left lane car is offscreen reset the picbox to the top
            ObLycord = 0
            PicLeft.Top = ObLycord
            PicLeft.Picture = LoadPicture()
            LeftCar = False
            
        End If
    End If
        
    'Check for colision with left lane car
    If Within(ObLxcord, ObLycord, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
    If Within(ObLxcord + PicLeft.Width, ObLycord, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
    If Within(ObLxcord, ObLycord + PicLeft.Height, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
    If Within(ObLxcord + PicLeft.Width, ObLycord + PicLeft.Height, xcord, ycord, xcord + picCar.Width, ycord + picCar.Height) = True Then gameover = True
    
    'loop for correct time of the next loop
    Do
        DoEvents
    Loop Until timeGetTime >= nTime
Loop Until gameover = True

    'take action if game has ben won
    If Winner = True Then
        Dim secondtaken As Integer
        secondtaken = 59 - second
        MsgBox "Congratulation you passed then " & Win & " cars goal" _
                & vbNewLine & "You have done it in " & TimeTaken & ":" & secondtaken
        NextWin = Win + 5
    Else    'taken avtion if game has been lost
        picCar.Picture = Image1.Picture
        MsgBox "GameOver, you have crashed" & vbNewLine & "Maybe you need more practice" & _
                vbNewLine & "Thx for playing"
    End If
    TimeTimer.Enabled = False
    Command1.Enabled = True
End Sub

Public Function Within(x1, y1, x2, y2, x3, y3) As Boolean
    'mathematical function to calculate a valid collision between 2 objects
    If x1 >= x2 And x1 <= x3 And y1 <= y3 And y1 >= y2 Then
        Within = True
    End If
End Function

Public Function PlaceCars()
Dim selectRightCar
    'if there is no right lane cars then place a right lane car
    If RightCar = False Then
        Randomize
        'rnd function which will randomly choose a car
        selectRightCar = (3 * Rnd)
        'set the car in the right lane picbox
        PicRight.Picture = imgRightCar(selectRightCar).Picture
        RightCar = True
    End If
Dim SelectLeftCar As Integer
Dim spacing As Integer
    'if there is no car in the left lane then place a left lane car
    If LeftCar = False Then
        Randomize
        'random function used to determine when a left lane car will show up
        'this will make it more challenging since you never know
        'when the next car will pop up
        spacing = Int(39 * Rnd) + 1
        If spacing = 20 Then
            Randomize
            'random function used to select a car
            SelectLeftCar = (3 * Rnd)
            'random function used to select a speed of the car
            'depending on a number the car will either come
            'foward fast or slow making it more harder
            LeftObjectSpeed = Int((25 * Rnd) + 5)
            'place the apropriate car in left lane picbox
            PicLeft.Picture = imgLeftCar(SelectLeftCar).Picture
            LeftCar = True
        End If
    End If
       
End Function


Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuThehelp_Click()
frmhelp.Show
End Sub

Private Sub TimeTimer_Timer()
'timer sub which will do a timer
'the timer always starts a 2minute and will decrease every second
    If second >= 0 Then
        If second > 9 Then
            txtTime.Text = minute & ":" & second
        Else
            txtTime.Text = minute & ":0" & second
        End If
    Else
        minute = minute - 1
        second = 59
        txtTime.Text = minute & ":" & second
    End If
    If minute < 0 Then
        txtTime.Enabled = False
        txtTime.Text = "0:00"
    End If
    If second = 0 Then
        TimeTaken = TimeTaken + 1
    End If
    second = second - 1
End Sub
