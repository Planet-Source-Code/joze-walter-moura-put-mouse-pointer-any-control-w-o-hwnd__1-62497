VERSION 5.00
Begin VB.Form frmTstPutMouse 
   Caption         =   " Putting Mouse over any Control"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Cbo1 
      Height          =   315
      Left            =   90
      TabIndex        =   17
      Text            =   "Cbo1"
      Top             =   2580
      Width           =   1440
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   1725
      TabIndex        =   16
      Top             =   4830
      Width           =   2385
   End
   Begin VB.Frame Pan1 
      Caption         =   "Panel"
      Height          =   900
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   1380
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "Co&MboBox"
      Height          =   330
      Index           =   7
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2025
      Width           =   2010
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&DirListBox"
      Height          =   330
      Index           =   6
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2025
      Width           =   2015
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&CheckBox"
      Height          =   330
      Index           =   5
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   2010
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&TextBox"
      Height          =   330
      Index           =   4
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   2015
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&Frame"
      Height          =   330
      Index           =   3
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1350
      Width           =   2010
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&Image"
      Height          =   330
      Index           =   2
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1350
      Width           =   2015
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&Label"
      Height          =   330
      Index           =   1
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1005
      Width           =   2010
   End
   Begin VB.CommandButton Cmd 
      Appearance      =   0  'Flat
      Caption         =   "&Buttons"
      Height          =   330
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1005
      Width           =   2015
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1740
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmTstPutMouse.frx":0000
      Top             =   3810
      Width           =   2370
   End
   Begin VB.CheckBox Ckb1 
      Caption         =   "Click to Check/Uncheck"
      Height          =   285
      Left            =   1740
      TabIndex        =   4
      Top             =   3480
      Width           =   2280
   End
   Begin VB.CommandButton Buts 
      Caption         =   "(3) Goto Botton 1"
      Height          =   495
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   4095
      Width           =   1425
   End
   Begin VB.CommandButton Buts 
      Caption         =   "(2) Goto Button 3"
      Height          =   495
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   3540
      Width           =   1425
   End
   Begin VB.CommandButton Buts 
      Caption         =   "(1) Goto Button 2"
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   2970
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Where PUT MOUSE Pointer?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   4050
   End
   Begin VB.Image Img1 
      Height          =   750
      Left            =   1740
      Picture         =   "frmTstPutMouse.frx":005D
      Stretch         =   -1  'True
      Top             =   2565
      Width           =   2355
   End
   Begin VB.Label Lab1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " THAT'S ALL FOLK! (Joze)"
      Height          =   195
      Left            =   2175
      TabIndex        =   3
      Top             =   5670
      Width           =   1905
   End
   Begin VB.Menu filFile 
      Caption         =   "&File"
      Begin VB.Menu filExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu pre 
      Caption         =   "&Preferences"
      Begin VB.Menu preCenter 
         Caption         =   "&Center the Mouse Pointer"
         Checked         =   -1  'True
      End
      Begin VB.Menu preAtLeft 
         Caption         =   "&Mouse At Left Region of Control"
      End
   End
   Begin VB.Menu tut 
      Caption         =   "&Tutorial"
      Begin VB.Menu tutStep1 
         Caption         =   "Step &1 Choose a Folder"
      End
      Begin VB.Menu tutStep2 
         Caption         =   "Step &2 Chek/Uncheck Your Option"
      End
      Begin VB.Menu tutStepn 
         Caption         =   "Step &n ... wold be: Type Your Name "
      End
   End
End
Attribute VB_Name = "frmTstPutMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'.----------------------------------------------------------------------
' Tst PutMouse : Put Mouse Pointer over any Control (no hWnd needed)
'==============:========================================================
' Author/Date  : J o z e - em 09/09/2005
' Description  : See modPutMouse.Bas, a few code lines to do it!
'              : All you paste your app: The .Bas so use the
'              : PutMouseOn and PutMouseAt to center or near left the
'              : pointer you design.
'              :
'              : All in this form are examples and code artefacts.
'`----------------------------------------------------------------------

Option Explicit

'This is the unique module reference
Private Sub PutMouse(ctl As Control)
   If preCenter.Checked Then
      PutMouseOn ctl, Me ' cursor is centered on control
   Else
      PutMouseAt ctl, Me ' cursor is left inside the control
   End If
End Sub

Private Sub Cmd_Click(Index As Integer)
   Select Case Index
     Case 0
        PutMouse Buts(0)
     Case 1
        PutMouse Lab1
     Case 2
        PutMouse Img1
     Case 3
        PutMouse Pan1
     Case 4
        PutMouse Txt1
     Case 5
        PutMouse Ckb1
     Case 6
        PutMouse Dir1
     Case 7
        PutMouse Cbo1
   End Select
End Sub

Private Sub Buts_Click(Index As Integer)
   PutMouse Buts((Index + 1) Mod 3)
End Sub


Private Sub filExit_Click()
   Unload Me
End Sub


Private Sub Alternate_Checked_Menu()
   preCenter.Checked = preAtLeft.Checked
   preAtLeft.Checked = Not preAtLeft.Checked
End Sub


Private Sub preCenter_Click()
   Alternate_Checked_Menu
End Sub

Private Sub preAtLeft_Click()
   Alternate_Checked_Menu
End Sub

'WOULD BE A KIND OF TUTORIAL, WIZARD, HELP, ETC
'==============================================
Private Sub tutStep1_Click()
   PutMouse Dir1
   tutStep1.Checked = True
End Sub

Private Sub tutStep2_Click()
   PutMouse Ckb1
   tutStep2.Checked = True
End Sub

Private Sub tutStepn_Click()
   PutMouse Txt1
   tutStepn.Checked = True
End Sub
