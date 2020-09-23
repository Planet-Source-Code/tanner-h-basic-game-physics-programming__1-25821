VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Physics Demo ©2001 Tanner Helland (tannerhelland.50megs.com)"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   480
      Picture         =   "Physics.frx":0000
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicBBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   240
      Picture         =   "Physics.frx":0074
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox PicScreenBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6900
      Left            =   1200
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   456
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   6900
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00000000&
      Height          =   6900
      Left            =   1200
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   456
      TabIndex        =   12
      Top             =   120
      Width           =   6900
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1020
      Left            =   240
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   360
      Picture         =   "Physics.frx":00E8
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   120
      Picture         =   "Physics.frx":015C
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PicTM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   8520
      Picture         =   "Physics.frx":01D0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   7920
      Picture         =   "Physics.frx":1F15
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicRM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   6840
      Picture         =   "Physics.frx":3DF8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   5760
      Picture         =   "Physics.frx":5826
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicLM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4680
      Picture         =   "Physics.frx":745B
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3600
      Picture         =   "Physics.frx":8EED
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   12060
      Left            =   480
      Picture         =   "Physics.frx":AAFA
      ScaleHeight     =   800
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   12060
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1200
      Picture         =   "Physics.frx":10969
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   120
      Picture         =   "Physics.frx":126AE
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VELOCITY"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   960
   End
   Begin VB.Label LblVel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9960
      TabIndex        =   13
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A Basic Physics Demo; ©2001 by Tanner Helland

'Another staple of game programming is basic physics coding.  Here's a velocity
'demo for both a primary object (the ship) and a secondary object (the bullets).
'Play around with firing while moving - it makes some crazy effects at times.
'The code is simple and well-commented.  Use it as you like but do not
'redistribute except in .exe format.

'E-mail questions or comments to tannerhelland@hotmail.com

'Look for additional great code samples at tannerhelland.50megs.com

Private Sub Command1_Click()
    GameActive = 0
   End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Depending on the direction that's being pressed, set the correct
'direction variable and the correct picture
If KeyCode = vbKeyLeft Then
    sLeft = 1
    Picture1.Picture = PicL.Picture
    Picture2.Picture = PicLM.Picture
End If
If KeyCode = vbKeyRight Then
    sRight = 1
    Picture1.Picture = PicR.Picture
    Picture2.Picture = PicRM.Picture
End If
If KeyCode = vbKeyUp Then sUp = 1
If KeyCode = vbKeyDown Then sDown = 1
'Fire the gun
If KeyCode = vbKeySpace Then Firing = 1
'End the game
If KeyCode = vbKeyEscape Then
    GameActive = 0
    'RestoreRes
    End
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Tell the program that the user no longer wants the ship to
'move in that direction...
If KeyCode = vbKeyLeft Then sLeft = 0
If KeyCode = vbKeyRight Then sRight = 0
If KeyCode = vbKeyUp Then sUp = 0
If KeyCode = vbKeyDown Then sDown = 0
'Fire the gun as necessary
If KeyCode = vbKeySpace Then Firing = 0
End Sub

Public Sub MainLoop()
    'The main loop...
    Dim TimeTaken As Single
    Do While BeginGame = 1
    TimeTaken = Timer
    'Do all of the game engine stuff...
    DrawStars
    VelocityCode
    FireBullets
    'Change the velocity caption
    LblVel.Caption = 100 - sVel & " kps"
    'Halt to check for outside events (key presses, etc.)
    DoEvents
    'Do it all again...
5     If Timer - TimeTaken < 0.025 Then GoTo 5
    Loop
End Sub

Private Sub Form_Load()
    InitializeGameEngine
    'When the form is shown, call the main loop
    BeginGame = 1
    MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
