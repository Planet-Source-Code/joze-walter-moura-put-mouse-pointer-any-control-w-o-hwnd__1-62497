Attribute VB_Name = "mPutMouse"
'.----------------------------------------------------------------------
' mPutMouse    : Mouse Putting any Control Subroutines hWnd independent
'==============:========================================================
' Author/Date  : J o z e - em 09/09/2005
'              :
' Description  : PutMouseOn - Center the cursor on Controls
'              : PutMouseAt - Cursor will be near left Control region.
'              :
'              : 2 Parameters: Control Name and Form (Me)
'              :
'              : So, any form in Project may code, eg,
'              :
'              :     PutMouseON Command1, Me
'              :
' Comments     : Well, I've seen analogous code that needs Control hWnd
'              : to get X, Y positions.
'              : So, I was thinking about a new kind of wizard/tutor for
'              : a project. Some controls without hWnd was failling Error.
'              :
'              : Then, I use the Form hWnd to get references and the Left
'              : ans Top properties of controls to determine X and Y values.
'              :
'              : This code will extent the Put Mouse resources to controls
'              : with those proporties, Left/Top, as Label, Image, Picture,
'              : etc. Not aplies to a few one, as Menu, StatBar.
'              :
'              : It supports any Screen resolution and Control expansions.
'              :
'              : But, if the control has transparent background without
'              : borders, I recommend use the PutMouseAt sub, because the
'              : expanded center may be not simetrycs at look.
'              :
'              : The unique difference in both subroutines is the Height
'              : measure twice used in place the Width one.
'              :
' License      : Freeware - you may distribute, alter, sold, anything
'              : as you want. Send me any enhances may be you made, ok?
'              :
'              : Enjoy,
'              : Joze.
'              : Rio de Janeiro, Brasil.
'              :
'`----------------------------------------------------------------------
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
'.-----------------------------------------------------------------------
'|PutMouseOn: Center Mouse Pointer
''=======================================================================
' example on form: PutMouseOn Command1, Me
'                  PutMouseOn Label1, Me
Public Sub PutMouseOn(ByVal ctl As Control, frm As Form)
   Dim pnt As POINTAPI
  With frm
   ClientToScreen .hwnd, pnt
   SetCursorPos _
      pnt.X + .ScaleX(ctl.Left, .ScaleMode, vbPixels) + .ScaleX(ctl.Width / 2, .ScaleMode, vbPixels), _
      pnt.Y + .ScaleY(ctl.Top, .ScaleMode, vbPixels) + .ScaleY(ctl.Height / 2, .ScaleMode, vbPixels)
  End With
End Sub

'.-----------------------------------------------------------------------
'|PutMouseAt: Mouse Pointer near left side
''=======================================================================
' example on form: PutMouseOn Command1, Me
'                  PutMouseOn Label1, Me
' Note: if you need reduce code space and no uses for this, remove it!
Public Sub PutMouseAt(ByVal ctl As Control, frm As Form)
   Dim pnt As POINTAPI
  With frm
   ClientToScreen .hwnd, pnt
   SetCursorPos _
      pnt.X + .ScaleX(ctl.Left, .ScaleMode, vbPixels) + .ScaleX(ctl.Height / 2, .ScaleMode, vbPixels), _
      pnt.Y + .ScaleY(ctl.Top, .ScaleMode, vbPixels) + .ScaleY(ctl.Height / 2, .ScaleMode, vbPixels)
  End With
End Sub

'-oOo-oOo-oOo-


