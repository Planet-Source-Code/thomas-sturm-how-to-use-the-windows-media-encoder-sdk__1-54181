VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Please wait ..."
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lblDummy 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait, loading Source-File ..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
