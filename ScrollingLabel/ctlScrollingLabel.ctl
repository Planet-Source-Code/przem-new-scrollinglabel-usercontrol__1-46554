VERSION 5.00
Begin VB.UserControl ctlScrollingLabel 
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   ScaleHeight     =   2565
   ScaleWidth      =   6585
   ToolboxBitmap   =   "ctlScrollingLabel.ctx":0000
   Begin VB.Timer tmrChkStatus 
      Interval        =   10
      Left            =   0
      Top             =   3000
   End
   Begin VB.Timer tmrScrolling 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3480
   End
   Begin VB.PictureBox picText 
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   120
      ScaleHeight     =   2370
      ScaleWidth      =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   5640
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ctlScrollLabel"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Tag             =   "5"
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image imgBottomScroll 
      Height          =   480
      Left            =   3720
      Picture         =   "ctlScrollingLabel.ctx":0314
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTopScroll 
      Height          =   480
      Left            =   3120
      Picture         =   "ctlScrollingLabel.ctx":061E
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNoBottom 
      Height          =   480
      Left            =   1320
      Picture         =   "ctlScrollingLabel.ctx":0928
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNoTop 
      Height          =   480
      Left            =   720
      Picture         =   "ctlScrollingLabel.ctx":0C32
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   2040
      Picture         =   "ctlScrollingLabel.ctx":0F3C
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBottom 
      Height          =   480
      Left            =   2640
      Picture         =   "ctlScrollingLabel.ctx":1246
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ctlScrollingLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================
'   UserControl:    ctlScrollingLabel             <[| (C) 2003 PRzEm |]>
'   Version:        1.0                           =======================
'   Author:         PRzEm
'   Date:           July 01, 2003
'   E-mail:         admin@videoclips.prv.pl
'========================================================================
'   This user control is used to add in your project a scrolling label.
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

Private booScrollingLabel_Top As Boolean
Private booScrollingLabel_Bottom As Boolean

Private intDirection As Integer
Private intFontHeight As Integer
Private mfonFont As StdFont
Private mpoiCursorPos As POINTAPI
Private Sub UserControl_Initialize()
'When the control is being initialized, arrows are not choosen.
booScrollingLabel_Top = False
booScrollingLabel_Bottom = False

'When the control is being initialized, show the appropriate arrows.
picTop.Picture = imgNoTop.Picture
picBottom.Picture = imgBottom.Picture

'When the control is being initialized, set the lblText's location.
With lblText
 .Top = 0
 .Left = 120
End With

End Sub
Private Sub UserControl_Resize()
'If the control is been decrease too much, rise up an error.
On Error GoTo Err

'Control's fit depending on objects' layout.
picText.Height = UserControl.Height - 220
picTop.Top = picText.Top
picBottom.Top = picText.Top + picText.Height - picBottom.Height
UserControl.Width = picTop.Left + picTop.Width + 120

'So the control is not decreased too much.
If (picTop.Top + picTop.Height) >= picBottom.Top Then
 picBottom.Top = 605
 picText.Height = 1025
 UserControl.Height = 1250
End If

Exit Sub
Err:
 picBottom.Top = 605
 picText.Height = 1025
 UserControl.Height = 1250

End Sub
Private Sub UserControl_InitProperties()
'If the control's container is in design mode, turn off the timer, which
'will cause the control to stop working.
tmrChkStatus.Enabled = Ambient.UserMode

'Default Caption property's value will be name given by her container.
Caption = Ambient.DisplayName

End Sub
Private Sub tmrChkStatus_Timer()
'This event takes place every 10 ms (interval=10) and applied to control the changes of
'the arrows status - choosen/not choosen.
Dim lonCStat As Long
Dim lonCurrhWnd As Long
Dim intLabelBottom As Integer

'Turn off the timer.
tmrChkStatus.Enabled = False

'Define the number that describe the label's bottom.
intLabelBottom = lblText.Top + lblText.Height

'With aid of two Windows API functions, define window handle, over the mouse button is.
lonCStat = GetCursorPos&(mpoiCursorPos)
lonCurrhWnd = WindowFromPoint(mpoiCursorPos.X, mpoiCursorPos.Y)

If booScrollingLabel_Top = False Then
'If the label is not on the top and mouse button is over the arrow, change
'the arrow's picture (scrooling started - imgTopScroll) and start scrolling label
'in appropriate direction.
 If lonCurrhWnd = picTop.hwnd And lblText.Top <> 0 Then
  booScrollingLabel_Top = True
  picTop.Picture = imgTopScroll.Picture
  intDirection = dirTop
  tmrScrolling.Enabled = True
 End If
Else
'If the label is not on the top and mouse button is no longer over the arrow, change
'the arrow's picture (scrooling could be continued - imgTop) and stop scrolling.
 If lonCurrhWnd <> picTop.hwnd And lblText.Top <> 0 Then
  booScrollingLabel_Top = False
  picTop.Picture = imgTop.Picture
  tmrScrolling.Enabled = False
'If the label is on the top and mouse button is no longer over the arrow, change
'the arrow's picture (scrooling not possible - imgNoTop) and and stop scrolling.
 ElseIf lonCurrhWnd <> picTop.hwnd And lblText.Top = 0 Then
  booScrollingLabel_Top = False
  picTop.Picture = imgNoTop.Picture
  tmrScrolling.Enabled = False
 End If
End If

If booScrollingLabel_Bottom = False Then
'If the label is not on the bottom and mouse button is over the arrow, change
'the arrow's picture (scrooling started - imgBottomScroll) and start scrolling label
'in appropriate direction.
 If lonCurrhWnd = picBottom.hwnd And picText.Height <= intLabelBottom Then
  booScrollingLabel_Bottom = True
  picBottom.Picture = imgBottomScroll.Picture
  intDirection = dirBottom
  tmrScrolling.Enabled = True
 End If
Else
'If the label is not on the bottom and mouse button is no longer over the arrow, change
'the arrow's picture (scrooling could be continued - imgBottom) and stop scrolling.
 If lonCurrhWnd <> picBottom.hwnd And picText.Height <= intLabelBottom Then
  booScrollingLabel_Bottom = False
  picBottom.Picture = imgBottom.Picture
  tmrScrolling.Enabled = False
'If the label is on the bottom and mouse button is no longer over the arrow, change
'the arrow's picture (scrooling not possible - imgNoBottom) and stop scrolling.
 ElseIf lonCurrhWnd <> picTop.hwnd And picText.Height >= intLabelBottom Then
  booScrollingLabel_Bottom = False
  picBottom.Picture = imgNoBottom.Picture
  tmrScrolling.Enabled = False
 End If
End If
'Turn on the timer.
tmrChkStatus.Enabled = True

End Sub
Private Sub tmrScrolling_Timer()
Dim intLabelBottom As Integer

'Define the number that describe bottom label.
intLabelBottom = lblText.Top + lblText.Height

'In case of in what direction label should be scrolled, take appropriate action.
Select Case intDirection
Case Is = dirTop
'While scrolling up, it is necessary to activate arrow of scrolling bottom.
 If picBottom.Picture <> imgBottom.Picture Then picBottom.Picture = imgBottom.Picture
'Scroll the label on specific height
 lblText.Top = lblText.Top + intFontHeight
'If the label gets to the top, change arrow and stop scrolling.
 If lblText.Top = 0 Then
  picTop.Picture = imgNoTop.Picture
  tmrScrolling.Enabled = False
 End If
Case Is = dirBottom
'While scrolling down, it is necessary to activate arrow of scrolling up.
 If picTop.Picture <> imgTop.Picture Then picTop.Picture = imgTop.Picture
 lblText.Top = lblText.Top - intFontHeight
End Select

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'If the control's container is in design mode, turn off the timer, which
'will cause the control to stop working.
tmrChkStatus.Enabled = Ambient.UserMode

'Get propertys from property bag.
BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
Enabled = PropBag.ReadProperty("Enabled", True)
ForeColor = PropBag.ReadProperty("ForeColor", &H8000000F)
ScrollingIntevral = PropBag.ReadProperty("ScrollingIntevral", 100)

Set Font = PropBag.ReadProperty("Font", mfonFont)

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'Save propertys in property bag.
PropBag.WriteProperty "BackColor", BackColor, &H8000000F
PropBag.WriteProperty "BorderStyle", BorderStyle, 0
PropBag.WriteProperty "Caption", Caption, Ambient.DisplayName
PropBag.WriteProperty "Enabled", Enabled, True
PropBag.WriteProperty "ForeColor", ForeColor, &H8000000F
PropBag.WriteProperty "ScrollingIntevral", ScrollingIntevral, 100

PropBag.WriteProperty "Font", Font, mfonFont

End Sub
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets text displayed on label."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
'Caption property come from Caption property's object's lblText value.
Caption = lblText.Caption

End Property
Public Property Let Caption(ByVal NewValue As String)
'New Caption value's is pass on to lblText object.

'Define font's height.
lblText.Caption = Mid(NewValue, 1, 1)
intFontHeight = lblText.Height

'Show new text in lblText and change value in property bag.
lblText.Caption = NewValue
UserControl.PropertyChanged "Caption"

'If picText border style's have to be decorate with border, must incresse
'a little his width - beauty issue ;)
If BorderStyle = [Fixed Single] Then
 picText.Width = lblText.Width + 320
Else
 picText.Width = lblText.Width + 240
End If

'Depending on text width, define arrows location.
picTop.Left = picText.Left + picText.Width + 240
picBottom.Left = picText.Left + picText.Width + 240

'Depending on text width and arrows location, define control's height and width.
UserControl.Height = picText.Top + picText.Height + 120
UserControl.Width = picTop.Left + picTop.Width + 120

End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets label's background color."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'BackColor property is stored in BackColor property picText object.
BackColor = picText.BackColor

End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
'New BackColor value is passed to picText object.
picText.BackColor = NewValue
UserControl.PropertyChanged "BackColor"

End Property
Public Property Get BorderStyle() As lbBorderStyleTypes
Attribute BorderStyle.VB_Description = "Returns/sets label's border style."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
'BorderStyle property is stored in BackColor property picText object.
BorderStyle = picText.BorderStyle

End Property
Public Property Let BorderStyle(ByVal NewValue As lbBorderStyleTypes)

'Be sure that attribute value to BorderStyle property is correct.
If NewValue = None Or NewValue = [Fixed Single] Then
'New BorderStyle value is passed to picText object.
 picText.BorderStyle = NewValue
 UserControl.PropertyChanged "BorderStyle"
Else
'Incorrect value BorderStyle property - show error message.
 Err.Raise Number:=vbObjectError + 32112, Description:="Nieprawid³owy parametr BorderStyle (tylko 0 lub 1)"
End If

End Property
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets label's font."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
'Font property is stored in Font property lblText object.
Set Font = lblText.Font

End Property
Public Property Set Font(ByVal NewValue As StdFont)
'New Font value is passed to lblText object.
On Error GoTo Err
Set lblText.Font = NewValue
UserControl.PropertyChanged "Font"

Exit Property
Err:
MsgBox "Read the error message in Set Font section", vbOKOnly + vbExclamation, "Scrollling Label"
'Probably you changed font name over the control. This operation is not allowed
'because an error occures ;( You can change the font name only when control is on
'new form. btw: if you know how to manage this error, please email me :)
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets label's text color."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'ForeColor property is stored in ForeColor property lblText object.
ForeColor = lblText.ForeColor

End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
'New ForeColor value is passed to lblText object.
lblText.ForeColor = NewValue
UserControl.PropertyChanged "ForeColor"

End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
'Enabled property is stored in Enabled property control.
Enabled = UserControl.Enabled

End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
'New Enabled value is passed to control object.
UserControl.Enabled = NewValue
UserControl.PropertyChanged "Enabled"

'Depending on control's condition - active or not - should modify object lblText and
'arrows on control.
Select Case NewValue
Case Is = True
 lblText.Enabled = True
 picTop.Picture = imgNoTop.Picture
 picBottom.Picture = imgBottom.Picture
Case Is = False
 lblText.Enabled = False
 picTop.Picture = imgNoTop.Picture
 picBottom.Picture = imgNoBottom.Picture
End Select

End Property
Public Property Get ScrollingIntevral() As Integer
Attribute ScrollingIntevral.VB_Description = "Returns/sets scrolling time."
Attribute ScrollingIntevral.VB_ProcData.VB_Invoke_Property = ";Behavior"
'Interval property is stored in Interval property tmrScrolling object.
ScrollingIntevral = tmrScrolling.Interval

End Property
Public Property Let ScrollingIntevral(ByVal NewValue As Integer)
'New ScrollingIntevral value is passed to tmrScrolling object.
tmrScrolling.Interval = NewValue
UserControl.PropertyChanged "ScrollingIntevral"

End Property
Public Sub DisplayAboutBox()
Attribute DisplayAboutBox.VB_Description = "Shows about box."
Attribute DisplayAboutBox.VB_UserMemId = -552
'Show About window.
frmAbout.Show vbModal
End Sub
