Attribute VB_Name = "ModResize"
'####
'#Start Po
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

