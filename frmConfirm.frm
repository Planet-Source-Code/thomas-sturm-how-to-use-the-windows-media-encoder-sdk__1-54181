VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6915
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraDummy 
      BorderStyle     =   0  'Kein
      Enabled         =   0   'False
      Height          =   2175
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   6975
      Begin VB.TextBox txtConfirm 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   8
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go !"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   0
      Left            =   -120
      TabIndex        =   0
      Top             =   -360
      Width           =   7125
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 4 : Check and confirm your Settings."
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
         Picture         =   "frmConfirm.frx":0CCA
         Top             =   480
         Width           =   2970
      End
   End
   Begin VB.Label lblDummy 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "If you want to change any of these Settings, please press the 'Back' - Button NOW !"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label lblDummy 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Please check and confirm your Settings."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Line lin3D 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6960
      Y1              =   4575
      Y2              =   4575
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
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmAddDetails.Show
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    frmDoIt.Show
    Me.Hide
End Sub

Private Sub Form_Activate()
    txtConfirm.Text = ""
    txtConfirm.Text = txtConfirm.Text & vbCrLf
    txtConfirm.Text = txtConfirm.Text & "   Source-Filename : " & SourceFileName & vbCrLf
    txtConfirm.Text = txtConfirm.Text & "   Output-Filename : " & DestFileName & vbCrLf & vbCrLf & vbCrLf
    txtConfirm.Text = txtConfirm.Text & "   Encoding-Profile : " & EncodingProfileName
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

