Attribute VB_Name = "Logic_Module"
'©2000-2001 Tanner Helland
Option Explicit
'ALL VARIABLES:

'Ship coordinates
Public ShipX As Integer, ShipY As Integer
'The ship's velocity in each direction
Public sLeft As Byte
Public sRight As Byte
Public sUp As Byte
Public sDown As Byte
Public sVelocity As Single
Public sVel As Single
'Whether or not the ship is firing
Public Firing As Byte
'The number of active bullets on the screen
Public BulletsActivated As Byte
Public Bullets(0 To NumOfBullets) As Bullet
'The array of stars
Public StarArray(0 To NumOfStars) As Star
'Whether or not to start the game
Public BeginGame As Byte
'Used to speed up the loops
Dim X As Integer, Y As Integer

Public Sub InitializeGameEngine()
'Randomize the random number generator
    Randomize
'Initially build all of the stars
    For X = 0 To NumOfStars
        BuildNewStar1 (X)
    Next X
'Set the ship in the middle of the picture box
    ShipX = Form1.PicMain.ScaleWidth / 2 - 32
    ShipY = Form1.PicMain.ScaleHeight - 64
    sVelocity = 10
'No active bullets
    BulletsActivated = 0
    BufferWidth = Form1.PicMain.ScaleWidth
    BufferHeight = Form1.PicMain.ScaleHeight
    Form1.Show
End Sub

Public Sub DrawStars()
'Temp variable to speed up the color determination
Static StarColor As Byte
Form1.PicScreenBuffer.Cls
'Draw the stars to their buffer
For X = 0 To NumOfStars
    StarColor = StarArray(X).bright
    StarArray(X).Y = StarArray(X).Y + StarArray(X).speed
    If StarArray(X).Y > BufferHeight Then BuildNewStar X
    SetPixelV Form1.PicScreenBuffer.hDC, StarArray(X).X, StarArray(X).Y, RGB(StarColor, StarColor, StarColor)
Next X

End Sub

Public Sub FireBullets()
'Run a loop through every bullet
For X = 0 To NumOfBullets
'If the user is firing, make a new bullet
If Firing = 1 And BulletsActivated <= 1 And Bullets(X).Activated = 0 Then
    Bullets(X).Activated = 1
    BulletsActivated = BulletsActivated + 1
    Bullets(X).X = ShipX + 30
    Bullets(X).Y = ShipY
End If
'If the bullet is active, do what you need to
If Bullets(X).Activated = 1 Then
    'Move the bullet up
    Bullets(X).Y = Bullets(X).Y - BulletSpeed
    'If it leaves the screen, deactivate it
    If Bullets(X).Y < -4 Then
        Bullets(X).Activated = 0
        GoTo 11
    End If
    'Blit the pictures to the buffer
    BitBlt Form1.PicScreenBuffer.hDC, Bullets(X).X, Bullets(X).Y, 4, 4, Form1.PicBulletM.hDC, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hDC, Bullets(X).X, Bullets(X).Y, 4, 4, Form1.PicBullet.hDC, 0, 0, vbSrcAnd
End If
'Next bullet
11 Next X
'Don't allow any more bullets to be created
BulletsActivated = 0
'Draw the entire buffer onto the screen (elimates flickering)
    BitBlt Form1.PicMain.hDC, 0, 0, BufferWidth, BufferHeight, Form1.PicScreenBuffer.hDC, 0, 0, vbSrcCopy
End Sub

Public Sub ScrollShip()
    'Scrolling across edges of screen
    If ShipX > Form1.PicMain.ScaleWidth Then ShipX = 0
    If ShipX < -64 Then ShipX = BufferWidth
    If ShipY > Form1.PicMain.ScaleHeight Then ShipY = 0
    If ShipY < -64 Then ShipY = BufferHeight
End Sub

Public Sub VelocityCode()
'Movement Up
If sUp = 1 Then
sVel = sVel - VAcc
ShipY = ShipY + sVel
End If

'Movement Down
If sDown = 1 Then
sVel = sVel + VAcc
ShipY = ShipY + sVel
End If

'Vertical Deceleration
If sUp = 0 And sDown = 0 And sVel <> 0 Then
If sVel > 0 Then
sVel = sVel - VDel
If sVel <= 0 Then sVel = 0
Else
sVel = sVel + VDel
If sVel >= 0 Then sVel = 0
End If
ShipY = ShipY + sVel
End If

'Movement Left
If sLeft = 1 Then
    sVelocity = sVelocity - HAcc
    ShipX = ShipX + sVelocity
End If

'Movement Right
If sRight = 1 Then
    sVelocity = sVelocity + HAcc
    ShipX = ShipX + sVelocity
End If

'If ShipX >= BufferWidth - 64 Then
'    ShipX = BufferWidth - 64
'    sVelocity = 0
'End If

'If ShipX <= 0 Then
'    ShipX = 0
'    sVelocity = 0
'End If


'Horizontal Deceleration
If sRight = 0 And sLeft = 0 And sVelocity <> 0 Then
    If sVelocity > 0 Then
        sVelocity = sVelocity - HDel
        If sVelocity <= 0 Then sVelocity = 0
    Else
        sVelocity = sVelocity + HDel
        If sVelocity >= 0 Then sVelocity = 0
    End If
    ShipX = ShipX + sVelocity
End If
'If the ship isn't moving, just draw it
If sVelocity = 0 Then
    Form1.Picture1.Picture = Form1.PicT.Picture
    Form1.Picture2.Picture = Form1.PicTM.Picture
End If
ScrollShip
'If you're not exploding, just draw the regular old ship
        BitBlt Form1.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, Form1.Picture2.hDC, 0, 0, vbMergePaint
        BitBlt Form1.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, Form1.Picture1.hDC, 0, 0, vbSrcAnd

End Sub

Public Sub BuildNewStar1(ByVal ArrayVal As Integer)
'Build a new star for the first time, setting random values
    StarArray(ArrayVal).X = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).Y = Rnd * Form1.PicMain.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 255
    StarArray(ArrayVal).speed = Rnd * 10 + 2
End Sub

Public Sub BuildNewStar(ByVal ArrayVal As Integer)
'Build a new star and make sure it's at the top of the screen
    StarArray(ArrayVal).X = Rnd * BufferWidth
    StarArray(ArrayVal).Y = 0
    StarArray(ArrayVal).bright = Rnd * 200 + 55
    StarArray(ArrayVal).speed = Rnd * 10 + 1
End Sub
