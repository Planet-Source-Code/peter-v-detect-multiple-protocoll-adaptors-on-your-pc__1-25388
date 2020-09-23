VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Find the IP addresses installed on your PC!!"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check IT !!"
      Height          =   555
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   3435
   End
   Begin VB.Label Label4 
      Caption         =   "Peter Verburgh."
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   3660
      TabIndex        =   5
      Top             =   2760
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Please Vote for me !             Look at my other code on PSC !!"
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Now you can use the right ip to use with winsock control (bind method !)"
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   3600
      TabIndex        =   3
      Top             =   1260
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Detect the IP addresses installed on your PC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Module1.Start
End Sub
