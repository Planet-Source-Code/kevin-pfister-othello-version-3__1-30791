VERSION 5.00
Begin VB.Form frmNames 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Names"
   ClientHeight    =   2265
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmNames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Computer 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox Player2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox player 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Computer"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Player Two"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Player One"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    player = GetSetting(App.EXEName, "PlayerID", "PlayerOne")
    Player2 = GetSetting(App.EXEName, "PlayerID", "PlayerTwo")
    Computer = GetSetting(App.EXEName, "PlayerID", "Computer")
End Sub

Private Sub OKButton_Click()
    SaveSetting App.EXEName, "PlayerID", "PlayerOne", player
    SaveSetting App.EXEName, "PlayerID", "PlayerTwo", Player2
    SaveSetting App.EXEName, "PlayerID", "Computer", Computer
    Unload Me
End Sub
