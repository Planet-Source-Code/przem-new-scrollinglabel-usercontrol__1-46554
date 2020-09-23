Attribute VB_Name = "modScrollingLabel"
'========================================================================
'   UserControl:    ctlScrollingLabel             <[| (C) 2003 PRzEm |]>
'   Version:        1.0                           =======================
'   Author:         PRzEm
'   Date:           July 01, 2003
'   E-mail:         admin@videoclips.prv.pl
'========================================================================
'   This usercontrol is used to add in your project a scrolling label.
'   User may create a several labels and put in it a lots of text.
'   Label will be scrolled down or up, until the whole text will be shown.
'
'   If you have more ideas and you like to add more functions to this user
'   control, go ahead ;) My programming skills aren't too big, but I hope
'   you will like this project. Please email me your modifications makes in
'   my user control.
'
'   This code may be reused and modified for non-commercial purposes only as
'   long as credit is given to the author in the programmes about box and
'   it's documentation.
'   If you use this code, please email me and let me know what you think about
'   this code and what you are doing with it.
'========================================================================
'   Bugs: if you find any bugs, email me at admin@videoclips.prv.pl
'========================================================================
Option Explicit

Type POINTAPI
 X As Long
 Y As Long
End Type
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function WindowFromPoint& Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long)

Public Enum Direction
 dirTop = 1
 dirBottom = 2
End Enum

Public Enum lbBorderStyleTypes
 None = 0
 [Fixed Single] = 1
End Enum

