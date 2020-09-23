VERSION 5.00
Begin VB.Form frmAddDetails 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmAddDetails.frx":0000
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
   Begin VB.TextBox txtRating 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   3840
      Width           =   5295
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox txtCopyrightInfo 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   5295
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   5295
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   5295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   8
      Top             =   -360
      Width           =   7125
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3 : Add your personal Details."
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
         TabIndex        =   9
         Top             =   480
         Width           =   3375
      End
      Begin VB.Image imgTopRight 
         Height          =   810
         Left            =   3960
         Picture         =   "frmAddDetails.frx":0CCA
         Top             =   480
         Width           =   2970
      End
   End
   Begin VB.Label lblDummy 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Note : You can leave these empty."
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   15
      Top             =   4200
      Width           =   5295
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Rating Information :"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   14
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Information :"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Author :"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Title :"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   3375
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
End
Attribute VB_Name = "frmAddDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmSelectProfile.Show
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    ContentTitle = txtTitle.Text
    AuthorName = txtAuthor.Text
    CopyrightInformation = txtCopyrightInfo.Text
    Description = txtDescription.Text
    RatingInformation = txtRating.Text
    frmConfirm.Show
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
