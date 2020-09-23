VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Reseider"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   0
      Left            =   9360
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   1
      Left            =   9600
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   2
      Left            =   9840
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   3
      Left            =   9840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   4
      Left            =   9840
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   5
      Left            =   9600
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   6
      Left            =   9360
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPatrate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   7
      Left            =   9360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command Resize"
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   5040
      ScaleHeight     =   1515
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label Resize"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tmpX As Long ' Temporary X
Public tmpy As Long ' Temporary Y
Public Obj_S As Control
Public MouseD_p As Boolean





Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Position_P Command1
Set Obj_S = Command1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Position_P Label1
Set Obj_S = Label1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Position_P Picture1
Set Obj_S = Picture1
End Sub



'####
'#Start Redimensionare obiect selectat
'############################
Private Sub picPatrate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  tmpX = X
  tmpy = Y
 MouseD_p = True

End Sub

Private Sub picPatrate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
  If Obj_S.Width < 100 Then
   Obj_S.Width = 101
   MouseD_p = False
   Exit Sub
   End If
  If Obj_S.Height < 100 Then
   Obj_S.Height = 101
   MouseD_p = False
   Exit Sub
  End If
  If MouseD_p = True Then
   'Colt Stanga sus
   If Index = 0 Then
    Obj_S.Move Obj_S.Left + (X - tmpX), Obj_S.Top + (Y - tmpy), Obj_S.Width - X + tmpX, Obj_S.Height - Y + tmpy
    picPatrate(0).Move picPatrate(0).Left + (X - tmpX), picPatrate(0).Top + (Y - tmpy)
    Position_P Obj_S, 0
   End If
   ' Mijloc sus
   If Index = 1 Then
   Obj_S.Move Obj_S.Left, Obj_S.Top + (Y - tmpy), Obj_S.Width, Obj_S.Height - Y + tmpy
   picPatrate(1).Move picPatrate(1).Left, picPatrate(1).Top + (Y - tmpy)
   Position_P Obj_S, 1
   End If
   'Colt Dreapta sus
   If Index = 2 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top + (Y - tmpy), Obj_S.Width + X - tmpX, Obj_S.Height - Y + tmpy
    picPatrate(2).Move picPatrate(2).Left + (X - tmpX), picPatrate(2).Top + (Y - tmpy)
   Position_P Obj_S, 2
   End If
   'Mijloc Dreapta
   If Index = 3 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width + X - tmpX, Obj_S.Height
    picPatrate(3).Move picPatrate(3).Left + (X - tmpX), picPatrate(3).Top
   Position_P Obj_S, 3
   End If
   'Colt dreapta jos
   If Index = 4 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width + X - tmpX, Obj_S.Height + Y - tmpy
    picPatrate(4).Move picPatrate(4).Left + (X - tmpX), picPatrate(4).Top + (Y - tmpy)
   Position_P Obj_S, 4
   End If
   'Mijloc jos
   If Index = 5 Then
    Obj_S.Move Obj_S.Left, Obj_S.Top, Obj_S.Width, Obj_S.Height + Y - tmpy
    picPatrate(5).Move picPatrate(5).Left, picPatrate(5).Top + (Y - tmpy)
   Position_P Obj_S, 5
   End If
   'Colt Stanga jos
   If Index = 6 Then
    Obj_S.Move Obj_S.Left + (X - tmpX), Obj_S.Top, Obj_S.Width - X + tmpX, Obj_S.Height + Y - tmpy
    picPatrate(6).Move picPatrate(6).Left + (X - tmpX), picPatrate(6).Top + (Y - tmpy)
    Position_P Obj_S, 6
   End If
   'Mijloc Jos
   If Index = 7 Then
    Obj_S.Move Obj_S.Left + (X - tmpX), Obj_S.Top, Obj_S.Width - X + tmpX, Obj_S.Height
    picPatrate(7).Move picPatrate(7).Left + (X - tmpX), picPatrate(7).Top
    Position_P Obj_S, 7
   End If
  End If

 
End Sub


Private Sub picPatrate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 MouseD_p = False
End Sub

'# Sfarsit Redimensionare obiect selectat
'############################

'+-------------------------------------------------------------------------------------+

'####
'#Start Pozitionare
'############################
Public Function Position_P(ByRef Obj As Control, Optional Ignorea As Integer = -1)
 
 
 
   With frmMain
    'Colt Stanga sus
    If Ignorea <> 0 Then
    .picPatrate(0).Move Obj.Left - 115, Obj.Top - 115
    .picPatrate(0).Visible = True
    End If
    'Mijloc sus
    If Ignorea <> 1 Then
    .picPatrate(1).Move Obj.Left + Obj.Width / 2 - 50, Obj.Top - 115
    .picPatrate(1).Visible = True
    End If
    'Colt Dreapta sus
    If Ignorea <> 2 Then
    .picPatrate(2).Move Obj.Left + Obj.Width + 15, Obj.Top - 115
    .picPatrate(2).Visible = True
    End If
    'Mijloc Dreapta
    If Ignorea <> 3 Then
    .picPatrate(3).Move Obj.Left + Obj.Width + 15, Obj.Top + Obj.Height / 2 - 50
    .picPatrate(3).Visible = True
    End If
    'Colt Dreapta jos
    If Ignorea <> 4 Then
    .picPatrate(4).Move Obj.Left + Obj.Width + 15, Obj.Top + Obj.Height + 15
    .picPatrate(4).Visible = True
    End If
    'Mijloc jos
    If Ignorea <> 5 Then
    .picPatrate(5).Move Obj.Left + Obj.Width / 2 - 50, Obj.Top + Obj.Height + 15
    .picPatrate(5).Visible = True
    End If
    'Colt Stanga jos
    If Ignorea <> 6 Then
    .picPatrate(6).Move Obj.Left - 115, Obj.Top + Obj.Height + 15
    .picPatrate(6).Visible = True
    End If
    'Mijloc Stanga
    If Ignorea <> 7 Then
    .picPatrate(7).Move Obj.Left - 115, Obj.Top + Obj.Height / 2 - 15
    .picPatrate(7).Visible = True
    End If
    
   End With
  
End Function


'# Sfarsit Start Drag Stanga
'############################

'+-------------------------------------------------------------------------------------+


