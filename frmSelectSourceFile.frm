VERSION 5.00
Begin VB.Form frmSelectSourceFile 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmSelectSourceFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6915
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output-File :"
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   6375
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   6135
      End
      Begin VB.CommandButton cmdBrowseOutput 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame fraInputFile 
      Caption         =   "Input-File :"
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   6375
      Begin VB.CommandButton cmdBrowseInput 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtSourceFile 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   7
      Top             =   -360
      Width           =   7125
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 : Please select Input- and Outputfile."
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
         TabIndex        =   8
         Top             =   480
         Width           =   3375
      End
      Begin VB.Image imgTopRight 
         Height          =   810
         Left            =   3960
         Picture         =   "frmSelectSourceFile.frx":0CCA
         Top             =   480
         Width           =   2970
      End
   End
   Begin VB.Line lin3D 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6960
      Y1              =   4580
      Y2              =   4580
   End
   Begin VB.Line lin3D 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "frmSelectSourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmMain.Show
    Me.Hide
End Sub

Private Sub cmdBrowseInput_Click()
    txtSourceFile.Text = OpenDialog(Me, "All Files (*.*)|*.*", "Select Source-File", "")
End Sub

Private Sub cmdBrowseOutput_Click()
    txtOutput.Text = SaveDialog(Me, "Windows Media Files (*.wmv)|*.wmv", "Select Output-File", "")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If txtSourceFile.Text = "" Or txtOutput.Text = "" Then
        MsgBox "Please set both Input AND Output-File !", vbCritical, "Error"
        Exit Sub
    End If
    Me.Hide
    frmWait.Show
    DoEvents
    SourceFileName = txtSourceFile.Text
    DestFileName = txtOutput.Text
    SetInputFile SourceFileName
    SetOutputFile DestFileName
    frmWait.Hide
    frmSelectProfile.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim retVal As Long
    retVal = MsgBox("Do you want to exit ?", vbQuestion + vbYesNo, "Exit ?")
    If retVal = vbYes Then
        CloseProgram
    Else
        Cancel = True
    End If
End Sub

