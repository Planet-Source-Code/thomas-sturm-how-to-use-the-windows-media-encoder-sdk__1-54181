VERSION 5.00
Begin VB.Form frmDoIt 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmDoIt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6915
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Timer tmrElapsed 
      Interval        =   1000
      Left            =   1200
      Top             =   4800
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   720
      Top             =   4800
   End
   Begin VB.Timer tmrPercent 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   4800
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Progress :"
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   6375
      Begin WMEnc.ProgressBar prgProgress 
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   661
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   777
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   0
      Top             =   -360
      Width           =   7125
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while the Encoder is working."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.Image imgTopRight 
         Height          =   810
         Left            =   3960
         Picture         =   "frmDoIt.frx":0CCA
         Top             =   480
         Width           =   2970
      End
   End
   Begin VB.Label lblRemaining 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblElapsed 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Time remaining :"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Time elapsed :"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Line lin3D 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line lin3D 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6960
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Label lblDest 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblSource 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Encoding Source-File"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmDoIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim retVal As Long
    retVal = MsgBox("Do you really want to abort ?", vbQuestion + vbYesNo, "Confirm")
    If retVal = vbYes Then
        Encoder.Stop
        frmFinish.Show
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    lblSource.Caption = SourceFileName
    lblDest.Caption = DestFileName
    Encode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = True
    End If
End Sub

Private Sub tmrPercent_Timer()
    Dim Stats As IWMEncStatistics
    Dim FileStats As IWMEncFileArchiveStats
    Dim lSize As WMENC_LONGLONG
    Dim lDuration As WMENC_LONGLONG
    Dim TempPercent As Variant
    Set Stats = Encoder.Statistics
    Set FileStats = Stats.FileArchiveStats
    lDuration = FileStats.FileDuration * 10000
    CurrentDuration = CDate(GetHHMMSS(lDuration))
    If CurrentDuration <> "00:00:00" Then
        TempPercent = Format((CurrentDuration / SourceDuration) * 100, "###")
        If TempPercent <> "" Then
            lblPercent.Caption = TempPercent & " %"
            prgProgress.Position = Val(TempPercent)
        End If
    End If
    iEncoderState = Encoder.RunState
    If iEncoderState = WMENC_ENCODER_STOPPED Then
        tmrPercent.Enabled = False
        tmrTime.Enabled = False
        tmrElapsed.Enabled = False
        frmFinish.Show
        Unload Me
    End If
End Sub

Private Sub Encode()
    Percentage = getPercentage(prgProgress.Position, prgProgress.Max)
    BeginProgress
    lblRemaining.Caption = TimeRemaining(Percentage)
    tmrPercent.Enabled = True
    tmrTime.Enabled = True
    bRunning = True
    Encoder.Start
End Sub

Private Sub tmrTime_Timer()
    Percentage = prgProgress.Position / prgProgress.Max * 100
    lblRemaining.Caption = TimeRemaining(Percentage)
End Sub

Private Sub tmrElapsed_Timer()
    TimeElapsed = TimeElapsed + 1000
    lblElapsed.Caption = GetHHMMSS(TimeElapsed)
End Sub

