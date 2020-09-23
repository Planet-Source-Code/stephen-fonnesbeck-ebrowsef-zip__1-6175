VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHBrowseForFolder Demo"
   ClientHeight    =   1440
   ClientLeft      =   3630
   ClientTop       =   3585
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7080
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse For Folder"
      Height          =   435
      Left            =   2543
      TabIndex        =   0
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Directory:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private getdir As String
'
'

Private Sub Command1_Click()
    
    getdir = BrowseForFolder(Me, "Select A Directory", Text1.Text)
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    Text1.Text = getdir
    
End Sub

Private Sub Form_Load()

  Text1.Text = CurDir

End Sub
