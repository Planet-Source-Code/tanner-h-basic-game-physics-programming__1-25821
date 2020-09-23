Attribute VB_Name = "Miscellaneous_Module"
'©2000-2001 Tanner Helland

'Handles for BitBlt and SetPixelV
Global MainHDC As Long
Global BufferHDC As Long

'ALL API CALLS:
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'ALL CONSTANTS:
Global Const NumOfStars As Byte = 250
Global Const BulletSpeed As Byte = 20
Global Const HAcc As Byte = 2
Global Const HDel As Byte = 2
Global Const VAcc As Byte = 2
Global Const VDel As Byte = 2
Global Const KeySpeed As Byte = 10
Global Const NumOfBullets As Byte = 50

'ALL TYPES:
Public Type Star
    X As Integer
    Y As Integer
    bright As Byte
    speed As Byte
End Type

Public Type Bullet
    X As Integer
    Y As Integer
    Velocity As Integer
    Activated As Byte
End Type

Public Type PointXY
    X As Integer
    Y As Integer
End Type

'ALL GLOBAL VARIABLES:
Global BufferWidth As Integer
Global BufferHeight As Integer
