VERSION 5.00
Begin VB.Form frmFinish 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmFinish.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Finish"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   7095
      Begin VB.Image imgSideLeft 
         Height          =   4665
         Left            =   120
         Picture         =   "frmFinish.frx":0CCA
         Top             =   0
         Width           =   2550
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish"
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
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "The Wizard has completed encoding the Source-File to Windows-Media Format."
         Height          =   1215
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Please click 'Finish' to exit."
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   2
         Top             =   4200
         Width           =   3735
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for using WMEnc !"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   1
         Top             =   3000
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
    CloseProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseProgram
End Sub
