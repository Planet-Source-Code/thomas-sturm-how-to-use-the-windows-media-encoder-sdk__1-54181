VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmMain.frx":0000
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   10
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
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   -120
      TabIndex        =   2
      Top             =   -120
      Width           =   7095
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "- Encode the File"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   11
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Please click 'Next' to continue."
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   9
         Top             =   4200
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "- Adding Information to the Windows-Media File"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   8
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "- Selecting an Encoding-Profile"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   7
         Top             =   3120
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "- Selecting a Source-File"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   6
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "These steps are :"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   5
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "This Wizard will guide you through the steps neccessary to convert a Video-File of your choice to the Windows-Media Format."
         Height          =   1215
         Index           =   1
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   3735
      End
      Begin VB.Image imgSideLeft 
         Height          =   4665
         Left            =   120
         Picture         =   "frmMain.frx":0CCA
         Top             =   0
         Width           =   2550
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    frmSelectSourceFile.Show
    Me.Hide
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
