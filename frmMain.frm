VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INI Tutorial"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open INI File"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save New"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraNew 
      Caption         =   "New Settings"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtNewData 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtNewTime 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtNewDate 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblNewDate 
         Caption         =   "New Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblNewTime 
         Caption         =   "New Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblNewData 
         Caption         =   "New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame fraCurrent 
      Caption         =   "Current Settings"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtOldData 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtOldTime 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtOldDate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblOldData 
         Caption         =   "Old Data"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblOldTime 
         Caption         =   "Old Time"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblOldDate 
         Caption         =   "Old Date"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Load ini file into "Current" frame
    Dim iniFile As String, ItsThere As Boolean
    iniFile = App.Path & "\settings.ini"
    ItsThere = FileExists(iniFile)
    If ItsThere = False Then
        Open iniFile For Output As #1
        Print #1, "[Program Settings]"
        Print #1, "TheDate = " & Date
        Print #1, "TheTime = " & Time
        Print #1, "TheData = This is just a sample."
        Close #1
        cmdRefresh_Click
    Else
        cmdRefresh_Click
    End If
    txtNewDate.Text = Date
    txtNewTime.Text = Time
    txtNewData.Text = "Put your text here!"
End Sub

Private Sub cmdRefresh_Click()
    Dim iniFile As String
    iniFile = App.Path & "\settings.ini"
    txtOldDate.Text = ReadINI("Program Settings", "TheDate", iniFile)
    txtOldTime.Text = ReadINI("Program Settings", "TheTime", iniFile)
    txtOldData.Text = ReadINI("Program Settings", "TheData", iniFile)
End Sub

Private Sub cmdSave_Click()
    Dim iniFile As String
    iniFile = App.Path & "\settings.ini"
    WriteINI "Program Settings", "TheDate", txtNewDate.Text, iniFile
    WriteINI "Program Settings", "TheTime", txtNewTime.Text, iniFile
    WriteINI "Program Settings", "TheData", txtNewData.Text, iniFile
    cmdRefresh_Click
    txtNewDate.Text = Date
    txtNewTime.Text = Time
    txtNewData.Text = "Put your text here!"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    frmShowINI.Show
End Sub

Private Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
    Open FullFileName For Input As #1
    Close #1
    FileExists = True
    Exit Function
MakeF:
    FileExists = False
End Function

Private Sub lblNewDate_DblClick()
    txtNewDate.Text = Date
End Sub

Private Sub lblNewTime_DblClick()
    txtNewTime.Text = Time
End Sub
