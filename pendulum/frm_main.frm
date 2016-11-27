VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Pendulum Motion"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14910
   FillStyle       =   0  'Solid
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer strt_tim 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   360
      Top             =   5280
   End
   Begin VB.Timer iniang_tmr 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   840
      Top             =   5280
   End
   Begin VB.Frame Frame2 
      Caption         =   " Motion Status  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   3975
      Begin VB.CommandButton stop_cmd 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "(degree / s) / s"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "degree / s"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label oscno_lbl 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Oscilations no. "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label acc_lbl 
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label vel_lbl 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Ang. Acceleration      ="
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Ang. Velocity             = "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label theta_lbl 
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Initial Values "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9600
      TabIndex        =   0
      Top             =   6480
      Width           =   5055
      Begin VB.TextBox mass_txt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         TabIndex        =   28
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Retard_txt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox strln_txt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox grav_txt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "9.81"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton start_cmd 
         Caption         =   "&Start"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton sub_cmd 
         Caption         =   "<<"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton add_cmd 
         Caption         =   ">>"
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "kg"
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Mass :"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "kg / s"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label retard_lbl 
         Caption         =   "Retarding cof. :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   " m"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "( m / s ) / s"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Str. Length :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Gravity :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label iniang_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label time_lbl 
      Caption         =   "0.000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Time  =               s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   530
      Width           =   135
   End
   Begin VB.Line string_ln 
      X1              =   7500
      X2              =   7500
      Y1              =   600
      Y2              =   6600
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addpressed As Boolean
Dim subpressed As Boolean

Const Pi = 3.1413

Private Sub add_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
addpressed = True
iniang_tmr.Enabled = True
start_cmd.Enabled = True
End If
End Sub

Private Sub add_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
addpressed = False
iniang_tmr.Enabled = False
End If
End Sub






Private Sub Form_Load()
addpressed = False
subpressed = False
Load frm_report
End Sub

Private Sub Form_paint()
Circle (string_ln.X2, string_ln.Y2), 150, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub iniang_tmr_Timer()
If (addpressed = True) And (iniang_lbl.Caption < 89) Then
    iniang_lbl.Caption = Val(iniang_lbl.Caption) + 1
    Call add(1)
End If
If (subpressed = True) And (iniang_lbl.Caption > 1) Then
    iniang_lbl.Caption = Val(iniang_lbl.Caption) - 1
    Call subt(1)
End If
End Sub



Private Sub start_cmd_Click()
If (Val(Retard_txt) < 0 Or Val(grav_txt) < 0) Then
    MsgBox "There can't be negative values for gravity or retarding coefficient.", vbCritical, "Sorry..."
    Exit Sub
End If
If (Val(strln_txt) <= 0 Or Val(mass_txt) <= 0) Then
    MsgBox "String and mass must be non-zero positive value.", vbCritical, "Sorry..."
    Exit Sub
End If
With start_cmd
If (.Caption = "&Start" Or .Caption = "&Continue") Then
    If (.Caption = "&Start") Then
        theta = (Pi / 180) * Val(iniang_lbl.Caption)
        frm_report.iniang_lbl.Caption = iniang_lbl.Caption
    End If
    Call strt(Val(grav_txt), Val(strln_txt), Val(Retard_txt), Val(mass_txt))
    strt_tim.Enabled = True
    .Caption = "&Pause"
    grav_txt.Enabled = False
    Retard_txt.Enabled = False
    strln_txt.Enabled = False
    add_cmd.Enabled = False
    sub_cmd.Enabled = False
    stop_cmd.Enabled = True
    mass_txt.Enabled = False
ElseIf .Caption = "&Pause" Then
    grav_txt.Enabled = True
    Retard_txt.Enabled = True
    strt_tim.Enabled = False
    .Caption = "&Continue"

End If
End With
End Sub

Private Sub stop_cmd_Click()
mass_txt.Enabled = True
add_cmd.Enabled = True
start_cmd.Enabled = False
sub_cmd.Enabled = True
strt_tim.Enabled = False
time_lbl.Caption = "0.000"
start_cmd.Caption = "&Start"
iniang_lbl.Caption = "0"
Retard_txt.Enabled = True
acc_lbl.Caption = ""
oscno_lbl.Caption = ""
vel_lbl.Caption = ""
grav_txt.Enabled = True
strln_txt.Enabled = True
Dim choice As Integer
choice = MsgBox("Would you like to have a final report about your experament ? ", vbYesNo, "Report")
If (choice = 6) Then
With frm_report
    .maxacc_lbl = Round(Abs(maxacc), 2)
    .maxvel_lbl = Round(Abs(maxvel), 2)
    .maxk_lbl = Round(0.5 * (maxvel * Val(frm_main.strln_txt)) ^ 2, 2)
    .maxp_lbl = Round(Val(frm_main.grav_txt) * Val(frm_main.strln_txt) * (1 - Cos(Val(.iniang_lbl) * (Pi / 180))), 2)
    .Show (1)
End With
End If
stop_cmd.Enabled = False
Cls
string_ln.Y2 = string_ln.Y1 + 6000
string_ln.X2 = string_ln.X1
Circle (string_ln.X2, string_ln.Y2), 150, 0
maxvel = 0
maxacc = 0


End Sub

Private Sub strt_tim_Timer()
Call strt(Val(grav_txt), Val(strln_txt), Val(Retard_txt), Val(mass_txt))
time_lbl.Caption = Val(time_lbl.Caption) + (strt_tim.Interval / 1000)
End Sub

Private Sub sub_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
subpressed = True
iniang_tmr.Enabled = True
End If
End Sub

Private Sub sub_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
subpressed = False
iniang_tmr.Enabled = False
End If
End Sub


