Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public xcord, ycord, ObLxcord, ObLycord, ObRxcord, ObRycord, objXcord, objYcord As Single
