VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find/Find Next/Replace Example"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Find/Find Next/Replace"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Replace"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find Next"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "the"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Replace With:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Created by Steve Bailey (http://www.xcellsoft.cjb.net"
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
FindMOD.FindAndHighlight Text1, Text2.Text, False
End Sub

Private Sub Command2_Click()
FindMOD.FindNextAndHighlight Text1, Text2.Text, False
End Sub

Private Sub Command3_Click()
FindMOD.ReplaceAndHighLight Text1, Text3.Text
End Sub
