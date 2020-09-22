VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please wait..."
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbProject 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2438
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbFile 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Module progress:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Project progress:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "fProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'fMain.Project.CancelEncryption
End Sub

