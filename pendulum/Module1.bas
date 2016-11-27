Attribute VB_Name = "Module1"
Option Explicit
Global theta   As Single
Global maxvel As Single
Global maxacc As Single
Dim prevstatusvel As Single




Public Function add(Disp As Single)
Const Pi = 3.1413
frm_main.Cls
Disp = (Pi / 180) * Disp
Dim prevang As Single
With frm_main.string_ln
    prevang = Atn((.X2 - .X1) / (.Y2 - .Y1))
    .Y2 = .Y2 - Sin(Disp / 2 + prevang) * 6000 * Sin(Disp)
    .X2 = .X2 + Cos(Disp / 2 + prevang) * 6000 * Sin(Disp)
    
    'Error correction due to uncertainity
If Sqr((.Y2 - .Y1) ^ 2 + (.X1 - .X2) ^ 2) > 6000 Then
Dim diff As Single
diff = Sqr((.Y2 - .Y1) ^ 2 + (.X1 - .X2) ^ 2) - 6000
.X2 = .X2 - diff * Sin(Disp + prevang)
.Y2 = .Y2 - diff * Cos(Disp + prevang)
End If

End With
frm_main.Circle (frm_main.string_ln.X2, frm_main.string_ln.Y2), 150, 0

End Function

Public Function subt(Disp As Single)
Const Pi = 3.1413
frm_main.Cls
Disp = (Pi / 180) * Disp
Dim prevang As Single
With frm_main.string_ln
    prevang = Atn((.X2 - .X1) / (.Y2 - .Y1))
    .Y2 = .Y2 + Sin(Disp / 2 + prevang) * 6000 * Sin(Disp)
    .X2 = .X2 - Cos(Disp / 2 + prevang) * 6000 * Sin(Disp)
     
'Error correction due to uncertainity

If Sqr((.Y2 - .Y1) ^ 2 + (.X1 - .X2) ^ 2) > 6000 Then
Dim diff As Single
diff = Sqr((.Y2 - .Y1) ^ 2 + (.X1 - .X2) ^ 2) - 6000
.X2 = .X2 - diff * Sin(Disp + prevang)
.Y2 = .Y2 - diff * Cos(Disp + prevang)
End If
End With
frm_main.Circle (frm_main.string_ln.X2, frm_main.string_ln.Y2), 150, 0


End Function

Public Sub strt(g As Single, l As Single, b As Single, m As Single)
Dim alpha As Single
Static omega As Single
Dim deltatheta As Single
If (frm_main.vel_lbl = "") Then omega = 0
Const Pi = 3.1413
With frm_main
alpha = -(1 / l) * g * Sin(theta) - b * omega / (m * l)
omega = omega + alpha * (.strt_tim.Interval / 1000)
deltatheta = omega * (.strt_tim.Interval / 1000)
If (deltatheta < 0) Then
    subt ((180 / Pi) * -deltatheta)
Else
    add ((180 / Pi) * deltatheta)
End If
.acc_lbl = Round((alpha * 180) / Pi, 2)
.vel_lbl = Round((omega * 180) / Pi, 2)
 theta = theta + deltatheta
 .iniang_lbl.Caption = Round(theta * (180 / Pi), 2)
If (Sgn(prevstatusvel) = 1 And Sgn(omega) = -1) Then
    .oscno_lbl.Caption = Val(.oscno_lbl.Caption) + 1
        frm_report.mp_lbl = Round(frm_main.time_lbl / .oscno_lbl.Caption, 2)
        frm_report.thp_lbl = Round(2 * Pi * Sqr(l / g), 2)
End If
If (Abs(maxvel) < Abs(omega)) Then
    maxvel = omega
End If
If (Abs(maxacc) < Abs(alpha)) Then
    maxacc = alpha
End If
End With
prevstatusvel = omega

End Sub
