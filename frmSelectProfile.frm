VERSION 5.00
Begin VB.Form frmSelectProfile 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "WMEnc"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmSelectProfile.frx":0000
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
   Begin VB.TextBox txtProfileDesc 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   9
      Top             =   3000
      Width           =   6615
   End
   Begin VB.CommandButton cmdPrfMgr 
      Caption         =   "Profile Manager"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
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
   Begin VB.ListBox lstProfiles 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Frame fraDummy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -120
      TabIndex        =   5
      Top             =   -360
      Width           =   7125
      Begin VB.Image imgTopRight 
         Height          =   810
         Left            =   3960
         Picture         =   "frmSelectProfile.frx":0CCA
         Top             =   480
         Width           =   2970
      End
      Begin VB.Label lblStep 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 : Please select an Encoding-Profile."
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
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Label lblDummy 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Use the Profile Manager to add, edit or delete Encoding-Profiles on your Computer."
      Height          =   975
      Index           =   2
      Left            =   4560
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Profile Description :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
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
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Profiles :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmSelectProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmSelectSourceFile.Show
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    EncodingProfileIndex = lstProfiles.ListIndex
    EncodingProfileName = ProColl.Item(EncodingProfileIndex).Name
    frmAddDetails.Show
    Me.Hide
End Sub

Private Sub cmdPrfMgr_Click()
    Dim TempVar As Integer
    TempVar = lstProfiles.ListIndex
    DisplayProfileManager
    EnumProfiles Me.lstProfiles
    lstProfiles.ListIndex = TempVar
End Sub

Private Sub Form_Load()
    EnumProfiles Me.lstProfiles
End Sub

Private Sub lstProfiles_Click()
    DisplayProfileDetails Me.txtProfileDesc, Me.lstProfiles, lstProfiles.ListIndex
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

