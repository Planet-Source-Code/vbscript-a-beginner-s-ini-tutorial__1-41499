VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShowINI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INI File"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbINIFile 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5530
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmShowINI.frx":0000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "frmShowINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    rtbINIFile.FileName = App.Path & "\settings.ini"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

