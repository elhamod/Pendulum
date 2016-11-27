VERSION 5.00
Begin VB.Form frm_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report about Motion"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " Energy Results "
      Height          =   4935
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "But... if they weren't equal... that is because a retarding force exists, consequently.... Energy 'LEAKS'..."
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label18 
         Caption         =   "Joule"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Joule"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "(rad / s) / s"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "rad / s"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "g * h ="
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "(1/2)v^2  ="
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "And their equality gives a solid evidence for the  Mechanical Energy conservation"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label maxp_lbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "max. Potential Energy per unit mass = "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label maxk_lbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "max. Kinetic Energy per unit  mass = "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "So... we can calculate :"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label maxacc_lbl 
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "max. ang. acc. ="
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label maxvel_lbl 
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "max. ang vel.  = "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   " by measurments : "
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Period Results "
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
      Begin VB.Label Label22 
         Caption         =   "s"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "s"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "If there is no period.... that's because the motion is either damped or overdamped..."
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label14 
         Caption         =   "degree "
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label iniang_lbl 
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Initial angle            ="
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label thp_lbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   $"frm_report.frx":0000
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Theoretical Period = "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label mp_lbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Measured Period   = "
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frm_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
Unload Me
Load frm_main
End Sub

