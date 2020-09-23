VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmFancontrol 
   AutoRedraw      =   -1  'True
   Caption         =   "PID PC Fan Control Demo"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   Icon            =   "frmFancontrol.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timerchart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   3240
   End
   Begin VB.CommandButton cmdplot 
      Caption         =   "Plot>"
      Height          =   375
      Left            =   4320
      TabIndex        =   41
      Top             =   4080
      Width           =   615
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   4815
      Left            =   5040
      OleObjectBlob   =   "frmFancontrol.frx":0442
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton cmdtp 
      Caption         =   "==>"
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      ToolTipText     =   "AMD Opteron Thermal Profile"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Timer Timersetp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   4680
   End
   Begin VB.CommandButton cmdprofile 
      Caption         =   "Set Thermal Profile"
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txterror 
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtkd 
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtki 
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtkp 
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Timer Timercpupower 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   4200
   End
   Begin VB.CommandButton cmddcpu 
      Caption         =   "Dynamic CPU Power"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Timer Timerfan 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   3720
   End
   Begin MSComctlLib.Slider sldsp 
      Height          =   600
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "could be set with Thermal Profile provide by CPU vendor"
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1058
      _Version        =   393216
      Max             =   100
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4200
      Top             =   3240
   End
   Begin MSComctlLib.Slider sldap 
      Height          =   600
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1058
      _Version        =   393216
      Enabled         =   0   'False
      Max             =   100
   End
   Begin MSComctlLib.Slider sldcpup 
      Height          =   600
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1058
      _Version        =   393216
      Min             =   3
      Max             =   100
      SelStart        =   3
      Value           =   3
   End
   Begin MSComctlLib.Slider sldfans 
      Height          =   600
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1058
      _Version        =   393216
      Enabled         =   0   'False
      Min             =   1200
      Max             =   4000
      SelStart        =   1200
      Value           =   1200
   End
   Begin MSComctlLib.Slider sldroom 
      Height          =   600
      Left            =   840
      TabIndex        =   8
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1058
      _Version        =   393216
      Max             =   40
   End
   Begin VB.Label Label11 
      Caption         =   "rpm"
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "W"
      Height          =   255
      Left            =   480
      TabIndex        =   38
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "'C"
      Height          =   255
      Left            =   480
      TabIndex        =   37
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lblrtdw 
      Caption         =   "'C"
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shapecpucore 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   5
      X1              =   6000
      X2              =   7440
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   4
      X1              =   6000
      X2              =   6000
      Y1              =   1440
      Y2              =   6000
   End
   Begin VB.Shape Shapefan 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   255
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   3
      X1              =   4200
      X2              =   4200
      Y1              =   2880
      Y2              =   5040
   End
   Begin VB.Shape Shapeap 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   2
      X1              =   4800
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape Shapeset 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   1
      X1              =   8280
      X2              =   8280
      Y1              =   960
      Y2              =   4080
   End
   Begin VB.Line Lineset 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   0
      X1              =   4800
      X2              =   8280
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shapeinlet 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10920
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Line Lineinlet2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   11040
      X2              =   11040
      Y1              =   240
      Y2              =   2760
   End
   Begin VB.Line Lineinlet 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   4800
      X2              =   11040
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   750
      Shape           =   3  'Circle
      Top             =   3270
      Width           =   465
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1590
      Top             =   3300
      Width           =   645
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   3270
      Width           =   465
   End
   Begin VB.Shape Shape17 
      FillColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1500
      Top             =   3960
      Width           =   795
   End
   Begin VB.Line Line9 
      X1              =   960
      X2              =   870
      Y1              =   3420
      Y2              =   3600
   End
   Begin VB.Line Line10 
      X1              =   960
      X2              =   1080
      Y1              =   3390
      Y2              =   3600
   End
   Begin VB.Line Line11 
      X1              =   870
      X2              =   1080
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line12 
      X1              =   750
      X2              =   420
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line13 
      X1              =   1230
      X2              =   1590
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line14 
      X1              =   2220
      X2              =   2610
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Label Label24 
      Caption         =   "X Gain"
      Height          =   225
      Left            =   1650
      TabIndex        =   33
      Top             =   3390
      Width           =   525
   End
   Begin VB.Label Label25 
      Caption         =   "Process"
      Height          =   225
      Left            =   1590
      TabIndex        =   32
      Top             =   4050
      Width           =   645
   End
   Begin VB.Line Line15 
      X1              =   990
      X2              =   990
      Y1              =   3750
      Y2              =   4170
   End
   Begin VB.Line Line16 
      X1              =   960
      X2              =   1470
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label26 
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1170
      TabIndex        =   31
      Top             =   3930
      Width           =   315
   End
   Begin VB.Label Label27 
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   30
      Top             =   3270
      Width           =   315
   End
   Begin VB.Line Line17 
      X1              =   750
      X2              =   660
      Y1              =   3510
      Y2              =   3450
   End
   Begin VB.Line Line18 
      X1              =   750
      X2              =   660
      Y1              =   3510
      Y2              =   3570
   End
   Begin VB.Line Line19 
      X1              =   990
      X2              =   930
      Y1              =   3750
      Y2              =   3840
   End
   Begin VB.Line Line20 
      X1              =   990
      X2              =   1050
      Y1              =   3750
      Y2              =   3840
   End
   Begin VB.Line Line23 
      X1              =   2610
      X2              =   2550
      Y1              =   3510
      Y2              =   3420
   End
   Begin VB.Line Line24 
      X1              =   2610
      X2              =   2550
      Y1              =   3510
      Y2              =   3600
   End
   Begin VB.Line Line21 
      X1              =   1590
      X2              =   1530
      Y1              =   3510
      Y2              =   3450
   End
   Begin VB.Line Line22 
      X1              =   1590
      X2              =   1530
      Y1              =   3510
      Y2              =   3600
   End
   Begin VB.Label Label28 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1320
      TabIndex        =   29
      Top             =   3300
      Width           =   195
   End
   Begin VB.Line Line25 
      X1              =   3060
      X2              =   3420
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line26 
      X1              =   3420
      X2              =   3420
      Y1              =   3510
      Y2              =   4170
   End
   Begin VB.Line Line27 
      X1              =   2280
      X2              =   3420
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Line Line28 
      X1              =   2280
      X2              =   2370
      Y1              =   4170
      Y2              =   4110
   End
   Begin VB.Line Line29 
      X1              =   2280
      X2              =   2370
      Y1              =   4170
      Y2              =   4230
   End
   Begin VB.Label Label31 
      Caption         =   "OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3540
      TabIndex        =   28
      Top             =   3750
      Width           =   405
   End
   Begin VB.Line Line30 
      X1              =   2940
      X2              =   2790
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Line Line31 
      X1              =   2790
      X2              =   2940
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line32 
      X1              =   2790
      X2              =   2880
      Y1              =   3420
      Y2              =   3510
   End
   Begin VB.Line Line33 
      X1              =   2880
      X2              =   2790
      Y1              =   3510
      Y2              =   3600
   End
   Begin VB.Label Label32 
      Caption         =   "Reset"
      Height          =   225
      Left            =   3090
      TabIndex        =   27
      Top             =   3240
      Width           =   465
   End
   Begin VB.Shape Shape18 
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2430
      Top             =   3750
      Width           =   225
   End
   Begin VB.Line Line34 
      X1              =   2640
      X2              =   2880
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line35 
      X1              =   2880
      X2              =   2880
      Y1              =   3870
      Y2              =   3750
   End
   Begin VB.Line Line36 
      X1              =   2880
      X2              =   2790
      Y1              =   3750
      Y2              =   3810
   End
   Begin VB.Line Line37 
      X1              =   2880
      X2              =   2970
      Y1              =   3750
      Y2              =   3810
   End
   Begin VB.Line Line38 
      X1              =   2490
      X2              =   2490
      Y1              =   3750
      Y2              =   3510
   End
   Begin VB.Label Label35 
      Caption         =   "d"
      Height          =   195
      Left            =   2460
      TabIndex        =   26
      Top             =   3780
      Width           =   135
   End
   Begin VB.Label Label36 
      Caption         =   "Rate"
      Height          =   165
      Left            =   2670
      TabIndex        =   25
      Top             =   3900
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Error"
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Derivitave Factor (rate)"
      Height          =   375
      Left            =   600
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Integral Factor (reset)"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Proportional Factor (gain)"
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblcpuc 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7780
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lblap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "22'C"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6690
      TabIndex        =   12
      Top             =   1395
      Width           =   405
   End
   Begin VB.Image imgfan3 
      Height          =   1800
      Left            =   3240
      Picture         =   "frmFancontrol.frx":2D1D
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image imgfan2 
      Height          =   1800
      Left            =   3240
      Picture         =   "frmFancontrol.frx":D263
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image imgfan1 
      Height          =   1800
      Left            =   3240
      Picture         =   "frmFancontrol.frx":17D91
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Label lblfs 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label lblcp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label lblsp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lblrt 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label 
      Caption         =   "Room Temperature"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Fan Speed"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "CPU Power"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Actual CPU Temperature"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Set CPU Temperature"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Image imgpc 
      Height          =   4695
      Left            =   5040
      Picture         =   "frmFancontrol.frx":2236B
      Stretch         =   -1  'True
      Top             =   15
      Width           =   6375
   End
   Begin VB.Image imgcpu 
      Height          =   1815
      Left            =   5160
      Picture         =   "frmFancontrol.frx":385B8
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   6375
   End
End
Attribute VB_Name = "frmFancontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim previous_error, error As Single
Dim integral As Single
Dim derivative As Single
Dim Kp, Ki, Kd As Single

Dim imagefan As Integer





Private Sub cmddcpu_Click()
If cmddcpu.Caption = "Dynamic CPU Power" Then

cmddcpu.Caption = "Constant CPU Power"
sldcpup.Value = Round(Rnd * 100)
Timercpupower.Enabled = True
Exit Sub
End If

If cmddcpu.Caption = "Constant CPU Power" Then
Timercpupower.Enabled = False
cmddcpu.Caption = "Dynamic CPU Power"
sldcpup.Value = 50
Exit Sub
End If
End Sub

Private Sub cmdplot_Click()
If cmdplot.Caption = "Plot>" Then
cmdplot.Caption = "<Plot"
MSChart.Enabled = True
MSChart.Visible = True
Timerchart.Enabled = True
Exit Sub
End If

If cmdplot.Caption = "<Plot" Then
cmdplot.Caption = "Plot>"
Timerchart.Enabled = False
MSChart.Enabled = False
MSChart.Visible = False
Exit Sub
End If



End Sub

Private Sub cmdprofile_Click()
If cmdprofile.Caption = "Set Thermal Profile" Then

cmdprofile.Caption = "Set CPU temperature"
sldsp.Enabled = False
Timersetp.Enabled = True
sldsp.Value = (sldcpup.Value / 3.864) + sldroom.Value + 7

Exit Sub
End If

If cmdprofile.Caption = "Set CPU temperature" Then
Timersetp.Enabled = False
cmdprofile.Caption = "Set Thermal Profile"
sldsp.Enabled = True
sldcpup.Value = 50
Exit Sub
End If

End Sub

Private Sub cmdtp_Click()
Load frmTP
frmTP.Show
End Sub

Private Sub Form_Load()
sldroom.Value = 25
sldsp.Value = 50
sldcpup.Value = 50
sldfans.Value = 1200
'===========================
'the below line was the hypothetical relationship of CPU temperature vs cpu power & cooling fan speed.
'this is not the real physical model. You need test & DOE to fit the real physical model.
sldap.Value = sldroom.Value + 10 * sldcpup.Value / sldcpup.Max * (sldfans.Max / sldfans.Value) + 25 * (sldcpup.Value / sldcpup.Max) ^ 2 * (sldfans.Max / sldfans.Value)
'===========================
lblrt.Caption = Str(sldroom.Value)
lblsp.Caption = Str(sldsp.Value)
lblap.Caption = Str(sldap.Value)
lblcp.Caption = Str(sldcpup.Value)
lblfs.Caption = Str(sldfans.Value)

previous_error = sldsp.Value - sldap.Value
integral = 0

Kp = -10 'proportional factor
Ki = -1 'integral factor
Kd = -10 'derivitave factor

txtkp.Text = Str(Kp)
txtki.Text = Str(Ki)
txtkd.Text = Str(Kd)
txterror.Text = Str(previous_error)

lblap.Left = imgpc.Left + 1580
lblap.Top = imgpc.Top + 1380

lblcpuc.BackColor = RGB(0, 0, 255)



Timer1.Enabled = True
Timerfan.Enabled = True
Timerfan.Interval = 20 * Round(sldfans.Max / sldfans.Value) 'fan speed animation control

'chart initial
Dim i As Integer
For i = 1 To 50
MSChart.Column = 1
MSChart.Row = i
MSChart.Data = 0
MSChart.Column = 2
MSChart.Row = i
MSChart.Data = 0
Next i

End Sub



Private Sub sldap_Change()

lblap.Caption = Str(sldap.Value)
'Timer1_Timer

End Sub

Private Sub sldcpup_Change()

lblcp.Caption = Str(sldcpup.Value)
'Timer1_Timer

End Sub

Private Sub sldfans_Change()

lblfs.Caption = Str(sldfans.Value)
Timerfan.Interval = 20 * Round(sldfans.Max / sldfans.Value)

'Timer1_Timer
End Sub

Private Sub sldroom_Change()
lblrt.Caption = Str(sldroom.Value)
'Timer1_Timer
End Sub


Private Sub sldsp_Change()

lblsp.Caption = Str(sldsp.Value)
'Timer1_Timer
End Sub

Private Sub Timer1_Timer()
error = sldsp.Value - sldap.Value
integral = integral + error * 1
derivative = (error - previous_error) / 1

sldfans.Value = sldfans.Value + Kp * error + Ki * integral + Kd * derivative

If sldfans.Value >= 4000 Then
sldfans.Value = 4000
integral = 0
End If
If sldfans.Value <= 1200 Then
sldfans.Value = 1200
integral = 0
End If
sldap.Value = sldroom.Value + 10 * sldcpup.Value / sldcpup.Max * (sldfans.Max / sldfans.Value) + 25 * (sldcpup.Value / sldcpup.Max) ^ 2 * (sldfans.Max / sldfans.Value)

previous_error = error

txterror.Text = Str(error)

Select Case sldap.Value
Case Is < 40
lblcpuc.BackColor = RGB(0, 0, 255)
Exit Sub

Case Is < 50
lblcpuc.BackColor = &HFFFF00
Exit Sub

Case Is < 60
lblcpuc.BackColor = &HFF00&
Exit Sub

Case Is < 70
lblcpuc.BackColor = &HFFFF&
Exit Sub

Case Is < 80
lblcpuc.BackColor = &H80FF&
Exit Sub

Case Is < 90
lblcpuc.BackColor = &HFF&
Exit Sub

Case Is < 100
lblcpuc.BackColor = &HFF&
Exit Sub

Case Else
Exit Sub

End Select


End Sub

Private Sub Timerchart_Timer()
Dim i As Integer
Dim nf As Single

MSChart.Column = 1
MSChart.Row = 50
MSChart.Data = sldsp.Value

MSChart.Column = 2
MSChart.Row = 50
MSChart.Data = sldap.Value

For i = 1 To 49
MSChart.Column = 2
MSChart.Row = i + 1
nf = MSChart.Data
MSChart.Row = i
MSChart.Data = nf

MSChart.Column = 1
MSChart.Row = i + 1
nf = MSChart.Data
MSChart.Row = i
MSChart.Data = nf



Next i



End Sub

Private Sub Timercpupower_Timer()
sldcpup.Value = Round(Rnd * 100)

End Sub

Private Sub Timerfan_Timer()
Select Case imagefan
Case 0
imgfan1.Visible = True
imgfan2.Visible = False
imgfan3.Visible = False
DoEvents
Case 1
imgfan1.Visible = False
imgfan2.Visible = True
imgfan3.Visible = False
DoEvents
Case 2
imgfan1.Visible = False
imgfan2.Visible = False
imgfan3.Visible = True
DoEvents
End Select
imagefan = imagefan + 1
If imagefan > 2 Then imagefan = 0
End Sub

Private Sub Timersetp_Timer()
sldsp.Value = (sldcpup.Value / 3.864) + sldroom.Value + 7

End Sub

Private Sub txtkd_Change()
Kd = Val(txtkd.Text)
End Sub

Private Sub txtki_Change()
Ki = Val(txtki.Text)
End Sub

Private Sub txtkp_Change()
Kp = Val(txtkp.Text)
End Sub
