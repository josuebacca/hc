Attribute VB_Name = "Module1"
Option Explicit

Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetCaretBlinkTime Lib "user32" () As Long
Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Global OriginalCaretBlinkTime%

