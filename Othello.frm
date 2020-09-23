VERSION 5.00
Begin VB.Form FrmOthello 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello"
   ClientHeight    =   6060
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9330
   Icon            =   "Othello.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   143
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   154
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   142
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   153
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   141
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   152
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   140
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   151
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   139
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   150
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   138
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   149
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   137
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   148
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   136
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   147
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   135
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   146
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   134
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   145
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   133
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   144
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   132
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   143
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   131
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   142
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   130
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   141
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   129
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   140
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   128
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   139
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   127
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   138
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   126
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   137
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   125
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   136
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   124
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   135
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   123
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   134
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   122
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   133
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   121
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   132
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   120
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   131
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   119
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   130
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   118
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   129
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   117
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   128
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   116
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   127
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   115
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   126
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   114
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   125
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   113
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   124
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   112
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   123
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   111
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   122
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   110
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   121
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   109
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   120
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   108
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   119
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   107
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   118
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   106
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   117
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   105
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   116
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   104
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   115
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   103
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   114
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   102
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   113
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   101
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   112
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   100
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   111
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   99
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   110
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   98
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   109
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   97
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   108
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   96
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   107
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   95
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   106
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   94
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   105
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   93
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   104
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   92
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   103
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   91
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   102
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   90
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   101
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   89
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   100
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   88
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   99
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   87
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   98
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   86
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   97
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   85
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   96
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   84
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   95
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   83
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   94
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   82
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   93
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   81
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   92
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   80
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   91
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   79
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   90
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   78
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   89
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   77
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   88
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   76
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   87
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   75
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   86
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   74
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   85
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   73
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   84
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   72
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   83
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   71
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   82
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   70
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   81
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   69
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   80
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   68
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   79
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   67
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   78
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   66
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   77
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   65
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   76
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   64
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   75
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   63
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   74
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   62
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   73
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   61
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   72
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   60
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   71
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   59
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   70
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   58
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   69
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   57
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   68
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   56
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   67
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   55
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   66
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   54
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   65
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   53
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   64
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   52
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   63
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   51
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   62
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   50
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   61
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   49
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   60
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   48
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   59
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   47
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   58
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   46
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   57
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   45
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   56
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   44
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   55
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   43
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   54
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   42
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   53
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   41
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   52
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   40
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   51
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   39
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   50
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   38
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   49
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   37
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   48
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   36
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   47
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   35
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   46
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   34
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   45
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   33
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   44
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   32
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   43
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   31
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   42
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   30
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   29
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   40
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   28
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   39
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   27
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   38
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   26
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   25
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   36
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   24
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   35
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   23
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   34
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   22
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   33
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   21
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   32
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   20
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   31
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   19
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   30
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   18
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   29
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   17
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   28
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   27
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   26
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   24
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   4920
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   21
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   3960
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   2040
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   1560
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      Appearance      =   0  'Flat
      BackColor       =   &H00249821&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox Lstprog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   2370
      Left            =   6000
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.PictureBox white 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      Picture         =   "Othello.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox black 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      Picture         =   "Othello.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblmove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label LblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label lblcom 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblplayer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lbltwo 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblone 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblgames 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Games = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnunew 
         Caption         =   "1 Player Game"
         Begin VB.Menu mnuX 
            Caption         =   "Play as Black"
            Checked         =   -1  'True
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuO 
            Caption         =   "Play as White"
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu mnunew2 
         Caption         =   "2 Player Game"
         Begin VB.Menu mnublack2 
            Caption         =   "Play as Black"
         End
         Begin VB.Menu mnuwhite2 
            Caption         =   "Play as White"
         End
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnushowp 
         Caption         =   "Show Possible Moves"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuskip 
         Caption         =   "Skip Go"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuNames 
         Caption         =   "Change Names"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuai 
      Caption         =   "AI"
      Begin VB.Menu mnucorn 
         Caption         =   "Smart, AI"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuview 
         Caption         =   "View Memory"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear AI Memory"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutip 
         Caption         =   "Tip of the Day"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "FrmOthello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Othello Created By Kevin Pfister
'This is a totally orginal version of Othello It includes excellent AI
'please vote or leave a comment about this or if
'there is a bug which i should know about.

'I know this is not the best coding espescially the number of nested loops, but i only took
'a small time making this. I hope you enjoy playing this.

'known Bugs
'   Sometimes the computer will repeatly choose the same sqaure
'   The computer may Enter a corner even if is already taken (Rare but is being Fixed)

'Array
Dim OthelloGrid(1 To 12, 1 To 12)         'This is the Grid on which the calcualations are done
Dim SelectableSquares(1 To 12, 1 To 12)   'Used to hold possible squares to move to
Dim checking(1 To 12, 1 To 12)            'Used to hold possible squares to move to
Dim ahead(1 To 12, 1 To 12)               'Used to hold possible user/computer moves

'Basic Stuff
Dim PlayerCounter As Long               'To determine who plays what
Dim goodshot As Long                    'To check if a shot is valid
Dim win As Long                         'To check if the shot was successful
Dim AIIsSmart As Long            'To determine if the corners are important to the AI
Dim oneplayer As Boolean                'has this game One Player
Dim squarepotent As Long                'This holds the Number of counters placed
Dim currentuser As Long
Dim ShowCMove As Long
Dim Playerone As String                 'This holds PlayerOne's Name
Dim playertwo As String                 'This holds PlayerTwo's Name
Dim Computer As String                  'This holds the Computer's name
Dim games As Long                       'This holds the number of games played

Private Sub Form_Load() 'Start of program
    Playerone$ = GetSetting(App.EXEName, "PlayerID", "PlayerOne", "PlayerOne")  'gets Player One's name from memory
    playertwo$ = GetSetting(App.EXEName, "PlayerID", "PlayerTwo", "PlayerTwo")  'gets Player Two's name from memory
    Computer$ = GetSetting(App.EXEName, "PlayerID", "Computer", "Computer")  'gets Computer's name from memory
    games = GetSetting(App.EXEName, "PlayerID", "NumberOfGames", 1)  'gets the number of Games from memory
    ShowCMove = 1
    revert 'Revert Colours Back to Normal
    currentuser = 1
    lblgames.Caption = games    'Sets the Games Label
    AIIsSmart = 1        'Sets Are corners important to True
    PlayerCounter = 1           'Play as Black
    oneplayer = True            'Sets the Game as one Player
    startup                     'Call the Startup sub
End Sub

Private Sub mnuabout_Click()    'The about menu option
    frmAbout.Show   'loads the about screeen
End Sub

Private Sub mnublack2_Click()   'The Two Player, First Player is Black New Game menu Option
    mnuX.Checked = False        'Not playing as Black
    mnuO.Checked = False        'Playing as White
    mnublack2.Checked = True    'Not playing as Black 2
    mnuwhite2.Checked = False   'Not playing as White 2
    mnuai.Enabled = False       'Disable AI
    
    PlayerCounter = 1           'Play as Black
    oneplayer = False           'Sets the Game as Two Player
    startup                     'Call the Startup sub
End Sub

Private Sub mnuclear_Click()    'Menu option to clear AI Memory
    SaveSetting App.EXEName, "AI", "AI Memory", ""  'Clears the AI Memory, setting
End Sub

Private Sub mnucorn_Click()     'Are corners important menu option
    If AIIsSmart Then    'Is it already checked
        mnucorn.Checked = False 'Uncheck the menu option
        AIIsSmart = 0    'Updates that corners are not important
    Else                        'Otherwise
        mnucorn.Checked = True  'Check the menu option
        AIIsSmart = 1    'Update that corners are important
    End If
End Sub

Private Sub mnuExit_Click() 'Exit menu option
    End                     'Exit the program
End Sub

Private Sub mnuNames_Click()    'Change the Names Menu option
    frmNames.Show               'Show the Name changing form
    Playerone$ = GetSetting(App.EXEName, "PlayerID", "PlayerOne")   'gets Player One's name from memory
    playertwo$ = GetSetting(App.EXEName, "PlayerID", "PlayerTwo")   'gets Player Two's name from memory
    Computer$ = GetSetting(App.EXEName, "PlayerID", "Computer")     'gets the Computer's name from memory
End Sub

Private Sub mnuO_Click()        'New game as white menu option
    mnuX.Checked = False        'Not playing as Black
    mnuO.Checked = True         'Playing as White
    mnublack2.Checked = False   'Not playing as Black 2
    mnuwhite2.Checked = False   'Not playing as White 2
    mnuai.Enabled = True        'AI is enabled
    
    PlayerCounter = 2           'Playas white
    oneplayer = True            'Sets the Game as one Player
    startup                     'Call the Startup sub
    comp (1)                    'Black(The computer) starts first
End Sub

Private Sub mnushowp_Click()
    If ShowCMove = 1 Then           'Is it already checked
        mnushowp.Checked = False    'Uncheck the menu option
        ShowCMove = 0               'Updates to not show possible Moves
        revert                      'Revert Colours Back to Normal
    Else                            'Otherwise
        mnushowp.Checked = True     'Check the menu option
        ShowCMove = 1               'Update to shows possible moves
        showmoves                   'Display Possible Moves
    End If
End Sub

Private Sub mnuskip_Click()                 'Skip go menu option
    If oneplayer = True Then                'If the game is oneplayer
        revert                              'Revert Colours Back to Normal
        comp (3 - PlayerCounter)            'Compute computers go
        If ShowCMove = 1 Then
            showmoves                       'Display Possible Moves
        End If
    Else                                    'OtherWise
        PlayerCounter = 3 - PlayerCounter   'Change the User
    End If
End Sub

Private Sub mnutip_Click()  'Tip of the Day menu option
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", 1   'Resets the Tips
    frmTip.Show             'Shows the Tips form
End Sub

Private Sub mnuview_Click() 'The AI Memory View menu option
    MsgBox (GetSetting(App.EXEName, "AI", "AI Memory")) 'Displays the Memory
End Sub

Private Sub mnuwhite2_Click()   'new Game as White 2 menu Option
    mnuX.Checked = False        'Not Playing as Black
    mnuO.Checked = False        'Not playing as White
    mnublack2.Checked = False   'Not playing as Black 2
    mnuwhite2.Checked = True    'Playing as White 2
    mnuai.Enabled = False       'Disable the AI Menu option
    
    PlayerCounter = 2           'Play as White
    oneplayer = False           'Sets the Game to 2 Player
    startup                     'Call the Startup sub
End Sub

Private Sub mnuX_Click()        'New game as black menu option
    mnuX.Checked = True         'Playing as Black
    mnuO.Checked = False        'Not playing as White
    mnublack2.Checked = False   'Not playing as Black 2
    mnuwhite2.Checked = False   'Not playing as White 2
    mnuai.Enabled = True        'Enables the AI menu option
    
    PlayerCounter = 1           'Play as Black
    oneplayer = True            'Sets the Game to 1 Player
    startup                     'Call the Startup sub
End Sub

Private Sub Othello_click(Index As Integer) 'User click grid for Users go
    FrmOthello.MousePointer = 11
    
    Xposition = Int(Index / 12) + 1  'Capture A cords
    Yposition = Index - (Xposition * 12) + 13 'Capture B cords
    If OthelloGrid(Xposition, Yposition) = 0 Then
        Call GameEngine(Xposition, Yposition, PlayerCounter, 2)   'Call placing sub
        Call message(Xposition & Yposition & squarepotent, False, False, False, False, True)
        If goodshot = 1 Then
            Call message(" valid Move", False, True, False, False, False) 'A valid move
            revert 'Revert Colours Back to Normal
            If oneplayer = True Then
                comp (3 - PlayerCounter)   'Compute computers go
                If ShowCMove = 1 Then
                    showmoves 'Display Possible Moves
                End If
            Else
                PlayerCounter = 3 - PlayerCounter
            End If
        Else
            Call message(" Invalid Move", False, True, False, True, False) 'Not a valid move
        End If
    Else
            Call message(" Invalid Move", False, True, False, True, False) 'Not a valid move
    End If
    checkwin    'call sub to check if there is a winner
    
    FrmOthello.MousePointer = 0
End Sub

Sub message(text$, stat As Boolean, move As Boolean, lst As Boolean, colour As Boolean, savetofile)
    If savetofile = True Then       'If save to file (Players movements only!)
        AIMemory$ = GetSetting(App.EXEName, "AI", "AI Memory")          'Get Settings
        SaveSetting App.EXEName, "AI", "AI Memory", AIMemory$ & text$   'Save settings + Extra
    End If
    If colour = True Then           'If colour is true then change the movement to Red
        lblmove.ForeColor = vbRed   'Change to Red
    Else                            'Otherwise
        lblmove.ForeColor = vbBlack 'Change to Black
    End If
    If move = True Then             'If movement is true
        lblmove.Caption = text$     'Display caption
    End If
    If stat = True Then             'If status is true
        LblStatus.Caption = text$   'Display status
    End If
    If lst = True Then              'If list is true
        Lstprog.AddItem (text$)     'Add to the List
    End If
End Sub


Sub startup()
    For outerloop = 1 To 12  'Clears the Grid Array
        For innerloop = 1 To 12
            OthelloGrid(outerloop, innerloop) = 0    'Clear grid
        Next
    Next
    
    lblone.Caption = " 2 Counter(s)" 'updates the lbls
    lbltwo.Caption = " 2 Counter(s)" 'updates the lbls
    lblgames.Caption = Val(lblgames.Caption) + 1    'Displays games this session
    SaveSetting App.EXEName, "PlayerID", "NumberOfGames", Val(lblgames.Caption)
    If PlayerCounter = 1 Then  'Updates the counter player lbls
        lblplayer.Caption = Playerone$ & " = Black"
        If oneplayer = True Then
            lblcom.Caption = Computer$ & " = White"
        Else
            lblcom.Caption = playertwo$ & " = White"
        End If
    Else
        lblplayer.Caption = Playerone$ & " = White"
        If oneplayer = True Then
            lblcom.Caption = Computer$ & " = Black"
        Else
            lblcom.Caption = playertwo$ & " = Black"
        End If
    End If
    Lstprog.Clear           'Clear computers captions
    OthelloGrid(6, 6) = 1   'Add a White
    OthelloGrid(7, 6) = 2   'Add a Black
    OthelloGrid(6, 7) = 2   'Add a Black
    OthelloGrid(7, 7) = 1   'Add a White
    For outerloop = 1 To 12  'This redraws the grid
        For innerloop = 1 To 12
            If OthelloGrid(outerloop, innerloop) = 0 Then
                Othello(((outerloop * 12) - 12) + innerloop - 1).Picture = LoadPicture("")
            ElseIf OthelloGrid(outerloop, innerloop) = 1 Then
                Othello(((outerloop * 12) - 12) + innerloop - 1).Picture = white.Picture
            ElseIf OthelloGrid(outerloop, innerloop) = 2 Then
                Othello(((outerloop * 12) - 12) + innerloop - 1).Picture = black.Picture
            End If
        Next
    Next
    Call message(" Status : " & Playerone$ & "'s Turn", True, False, False, False, False)  'Change game status
    If currentuser = 1 And ShowCMove = 1 Then
        showmoves 'Display Possible Moves
    End If

End Sub

Sub showmoves()
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            SelectableSquares(outerloop, innerloop) = 0  'Clears the SelectableSquares array
        Next
    Next
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If OthelloGrid(outerloop, innerloop) = 0 Then
                Call GameEngine(outerloop, innerloop, currentuser, 1)    'calls the othello game engine to check square
            End If
        Next
    Next
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If SelectableSquares(outerloop, innerloop) <> 0 Then
                Othello(((outerloop * 12) - 12) + innerloop - 1).BackColor = &HC0FFC0
            End If
        Next
    Next
End Sub

Sub revert()
    For revertloop = 0 To 143
        If Othello(revertloop).BackColor <> &H249821 Then
            Othello(revertloop).BackColor = &H249821
        End If
    Next
End Sub

Sub comp(currentuser)
    Call message(" Status : " & Computer$ & "'s Turn", True, False, True, False, False) 'Change game status
    
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            SelectableSquares(outerloop, innerloop) = 0  'Clears the SelectableSquares array
        Next
    Next
    
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If OthelloGrid(outerloop, innerloop) = 0 Then
                Call GameEngine(outerloop, innerloop, currentuser, 1)    'calls the othello game engine to check square
            End If
        Next
    Next
    
    bestxpos = 1    'Set Default
    bestypos = 1    'Set Default
    Bestscore = 0   'Set Default
    Call message(" Status : " & Computer$ & " Thinking...", True, False, True, False, False) 'Computer is checking possible moves
    
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If AIIsSmart Then
                If outerloop = 1 And innerloop = 1 And SelectableSquares(outerloop, innerloop) > 0 Then
                    bestxpos = outerloop
                    bestypos = innerloop
                    Bestscore = 100 + SelectableSquares(outerloop, innerloop)
                ElseIf outerloop = 1 And innerloop = 12 And SelectableSquares(outerloop, innerloop) > 0 Then
                    bestxpos = outerloop
                    bestypos = innerloop
                    Bestscore = 100 + SelectableSquares(outerloop, innerloop)
                ElseIf outerloop = 12 And innerloop = 1 And SelectableSquares(outerloop, innerloop) > 0 Then
                    bestxpos = outerloop
                    bestypos = innerloop
                    Bestscore = 100 + SelectableSquares(outerloop, innerloop)
                ElseIf outerloop = 12 And innerloop = 12 And SelectableSquares(outerloop, innerloop) > 0 Then
                    bestxpos = outerloop
                    bestypos = innerloop
                    Bestscore = 100 + SelectableSquares(outerloop, innerloop)
                End If
                If outerloop = 1 And innerloop = 2 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 1 And innerloop = 11 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 2 And innerloop = 1 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 2 And innerloop = 2 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 2 And innerloop = 11 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 2 And innerloop = 12 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 11 And innerloop = 1 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 11 And innerloop = 2 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 11 And innerloop = 11 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 11 And innerloop = 12 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 12 And innerloop = 2 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                ElseIf outerloop = 12 And innerloop = 11 And SelectableSquares(outerloop, innerloop) > 0 Then
                    SelectableSquares(outerloop, innerloop) = 1
                End If
            End If
            If SelectableSquares(outerloop, innerloop) > Bestscore Then
                bestxpos = outerloop
                bestypos = innerloop
                clearbrain
                brain = 1
                Bestscore = SelectableSquares(outerloop, innerloop)
            End If
            If SelectableSquares(outerloop, innerloop) = Bestscore And AIIsSmart Then
                If outerloop = 1 And innerloop = 2 Then
                    'Discard Before Corner
                ElseIf outerloop = 1 And innerloop = 11 Then
                    'Discard Before Corner
                ElseIf outerloop = 2 And innerloop = 1 Then
                    'Discard Before Corner
                ElseIf outerloop = 2 And innerloop = 2 Then
                    'Discard Before Corner
                ElseIf outerloop = 2 And innerloop = 11 Then
                    'Discard Before Corner
                ElseIf outerloop = 2 And innerloop = 12 Then
                    'Discard Before Corner
                ElseIf outerloop = 11 And innerloop = 1 Then
                    'Discard Before Corner
                ElseIf outerloop = 11 And innerloop = 2 Then
                    'Discard Before Corner
                ElseIf outerloop = 11 And innerloop = 11 Then
                    'Discard Before Corner
                ElseIf outerloop = 11 And innerloop = 12 Then
                    'Discard Before Corner
                ElseIf outerloop = 12 And innerloop = 2 Then
                    'Discard Before Corner
                ElseIf outerloop = 12 And innerloop = 11 Then
                    'Discard Before Corner
                Else
                    bestxpos = outerloop
                    bestypos = innerloop
                    ahead(bestxpos, bestypos) = 1
                    brain = brain + 1
                End If
            End If
        Next
    Next
    
    If bestxpos = 1 And bestypos = 1 And Bestscore = 0 Then
        Beep
        Call message("Unable to find Move", False, False, True, False, False) 'add this to the computer status list
        checkwin 'call sub to check if there is a winner
    Else
        If brain <> 1 And AIIsSmart Then
            winningscore = 0
            text$ = GetSetting(App.EXEName, "AI", "AI Memory", "111")
            For a = 1 To Len(text$) / 3
                Inner$ = Mid$(text$, a * 3, 3)
                xstr = Val(Mid$(Inner$, 1, 1))
                ystr = Val(Mid$(Inner$, 2, 1))
                Counter = Val(Mid$(Inner$, 3, 1))
                For outerloop = 1 To 12
                    For innerloop = 1 To 12
                        If ahead(outerloop, innerloop) <> 0 Then
                            If Counter > winningscore Then
                                bestxpos = outerloop
                                bestypos = innerloop
                                winningscore = Counter
                            End If
                        End If
                    Next
                Next
            Next
        End If
        clearbrain
        Call message(" Chosen Move " + Str$(bestxpos) + "," + Str$(bestypos), False, True, True, False, False) 'add this to the computer status list
        Call GameEngine(bestxpos, bestypos, currentuser, 2)    'Call the Othello game engine to place counter
        Call message(" Status : " & Playerone$ & "'s Turn", True, False, False, False, False) 'Change the game status
        checkwin 'call sub to check if there is a winner
    End If
End Sub

Sub GameEngine(EngineOuterLoop, EngineInnerLoop, currentuser, ClearOrPlace)    'This is the othello game engine

    'This Engine will check a square for possible square
    'Or it will be sent the Co-Ords of a square and will plot counters
    
    'This works by a recursive search
    
    'Say you start off with a square 4,2
    
    '* = You
    '+ = Enemy
    '# = Empty space
    
    ' # # # #
    ' # # # #
    ' # # + #
    ' # + # #
    ' * # # #
    
    'The computer will search around it and find any possible enemy squares
    'in this case 3,3
    'It will carry on in this direction until it finds another counter of the starting counter
    'It will SelectableSquares the Co-Ords of this square in an Array
    
    'It then can draw the pictures and update the grid
    
    ' # # # #
    ' # # # *
    ' # # * #
    ' # * # #
    ' * # # #
    
    'In this Engine there are nine different searches and they go in this order
    
    ' 1 2 3
    ' 4   5
    ' 6 11 12
    
    'The middle is left out because there is no need for this

    total = 0
    win = 0
    clearcheck
    If EngineOuterLoop - 1 > 0 And EngineInnerLoop - 1 > 0 Then
        If OthelloGrid(EngineOuterLoop - 1, EngineInnerLoop - 1) = currentuser Then
        
                
                crossed = 1
                If ClearOrPlace = 2 Then
                    checking(EngineOuterLoop - 1, EngineInnerLoop - 1) = 1
                End If
                
            For found = 1 To 12
                If EngineOuterLoop - (1 + found) > 0 And EngineInnerLoop - (1 + found) > 0 Then
                    If OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop - (1 + found)) = currentuser Then
                            crossed = crossed + 1
                            If ClearOrPlace = 2 Then
                                checking(EngineOuterLoop - (1 + found), EngineInnerLoop - (1 + found)) = 1
                            End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop - (1 + found)) = 3 - currentuser Then
                            total = total + crossed
                            If ClearOrPlace = 2 Then
                                Call drawpic(currentuser)
                            End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
    
    If EngineOuterLoop > 0 And EngineInnerLoop - 1 > 0 Then
        If OthelloGrid(EngineOuterLoop, EngineInnerLoop - 1) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop, EngineInnerLoop - 1) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop > 0 And EngineInnerLoop - (1 + found) > 0 Then
                    If OthelloGrid(EngineOuterLoop, EngineInnerLoop - (1 + found)) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop, EngineInnerLoop - (1 + found)) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop, EngineInnerLoop - (1 + found)) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop, EngineInnerLoop - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
    
    If EngineOuterLoop + 1 <= 12 And EngineInnerLoop - 1 > 0 Then
        If OthelloGrid(EngineOuterLoop + 1, EngineInnerLoop - 1) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop + 1, EngineInnerLoop - 1) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop + (1 + found) <= 12 And EngineInnerLoop - (1 + found) > 0 Then
                    If OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop - (1 + found)) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop + (1 + found), EngineInnerLoop - (1 + found)) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop - (1 + found)) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
    
    If EngineOuterLoop - 1 > 0 And EngineInnerLoop > 0 Then
        If OthelloGrid(EngineOuterLoop - 1, EngineInnerLoop) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop - 1, EngineInnerLoop) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop - (1 + found) > 0 And EngineInnerLoop > 0 Then
                    If OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop - (1 + found), EngineInnerLoop) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
     
    If EngineOuterLoop + 1 <= 12 And EngineInnerLoop > 0 Then
        If OthelloGrid(EngineOuterLoop + 1, EngineInnerLoop) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop + 1, EngineInnerLoop) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop + (1 + found) <= 12 And EngineInnerLoop > 0 Then
                    If OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop + (1 + found), EngineInnerLoop) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
     
    If EngineOuterLoop - 1 > 0 And EngineInnerLoop + 1 <= 12 Then
        If OthelloGrid(EngineOuterLoop - 1, EngineInnerLoop + 1) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop - 1, EngineInnerLoop + 1) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop - (1 + found) > 0 And EngineInnerLoop + (1 + found) <= 12 Then
                    If OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop + (1 + found)) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop - (1 + found), EngineInnerLoop + (1 + found)) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop + (1 + found)) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop - (1 + found), EngineInnerLoop + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
     
    If EngineOuterLoop > 0 And EngineInnerLoop + 1 <= 12 Then
        If OthelloGrid(EngineOuterLoop, EngineInnerLoop + 1) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop, EngineInnerLoop + 1) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop > 0 And EngineInnerLoop + (1 + found) <= 12 Then
                    If OthelloGrid(EngineOuterLoop, EngineInnerLoop + (1 + found)) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop, EngineInnerLoop + (1 + found)) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop, EngineInnerLoop + (1 + found)) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop, EngineInnerLoop + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If ClearOrPlace <> 1 Then
        clearcheck
    End If
     
    If EngineOuterLoop + 1 <= 12 And EngineInnerLoop + 1 <= 12 Then
        If OthelloGrid(EngineOuterLoop + 1, EngineInnerLoop + 1) = currentuser Then
        
            crossed = 1
            If ClearOrPlace = 2 Then
                checking(EngineOuterLoop + 1, EngineInnerLoop + 1) = 1
            End If
            
            For found = 1 To 12
                If EngineOuterLoop + (1 + found) <= 12 And EngineInnerLoop + (1 + found) <= 12 Then
                    If OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop + (1 + found)) = currentuser Then
                        crossed = crossed + 1
                        If ClearOrPlace = 2 Then
                            checking(EngineOuterLoop + (1 + found), EngineInnerLoop + (1 + found)) = 1
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop + (1 + found)) = 3 - currentuser Then
                        total = total + crossed
                        If ClearOrPlace = 2 Then
                            Call drawpic(currentuser)
                        End If
                    ElseIf OthelloGrid(EngineOuterLoop + (1 + found), EngineInnerLoop + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    If ClearOrPlace = 1 Then
        SelectableSquares(EngineOuterLoop, EngineInnerLoop) = total
    Else
        clearcheck
        squarepotent = total
        If win = 1 Then
            checking(EngineOuterLoop, EngineInnerLoop) = 1
            If currentuser = 1 Then
                Othello(((EngineOuterLoop * 12) - 12) + EngineInnerLoop - 1).Picture = black.Picture
                OthelloGrid(EngineOuterLoop, EngineInnerLoop) = 2
            Else
                Othello(((EngineOuterLoop * 12) - 12) + EngineInnerLoop - 1).Picture = white.Picture
                OthelloGrid(EngineOuterLoop, EngineInnerLoop) = 1
            End If
            goodshot = 1
        ElseIf win = 0 Then
            goodshot = 0
        End If
    End If
End Sub

Sub clearcheck()    'This clears the check array
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If checking(outerloop, innerloop) = 1 Then
                checking(outerloop, innerloop) = 0  'Sets the array to 0
            End If
        Next
    Next
End Sub

Sub clearbrain()    'This clears the brain array
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If ahead(outerloop, innerloop) = 1 Then
                ahead(outerloop, innerloop) = 0     'Sets the array to 0
            End If
        Next
    Next
End Sub

Sub drawpic(currentuser) 'Draws the picture to the grid referneces in the checking array
    win = 1
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If checking(outerloop, innerloop) = 1 Then
                If currentuser = 1 Then
                    Othello(((outerloop * 12) - 12) + innerloop - 1).Picture = black.Picture
                    OthelloGrid(outerloop, innerloop) = 2
                Else
                    Othello(((outerloop * 12) - 12) + innerloop - 1).Picture = white.Picture
                    OthelloGrid(outerloop, innerloop) = 1
                End If
            End If
        Next
    Next
End Sub

Sub checkwin()
    For outerloop = 1 To 12
        For innerloop = 1 To 12
            If OthelloGrid(outerloop, innerloop) = 1 Then
                NumberOfBlack = NumberOfBlack + 1
            ElseIf OthelloGrid(outerloop, innerloop) = 2 Then
                NumberOfWhite = NumberOfWhite + 1
            End If
        Next
    Next
    lblone.Caption = Str$(NumberOfBlack) + " Counter(s)" 'change the score bOthelloGrid
    lbltwo.Caption = Str$(NumberOfWhite) + " Counter(s)" 'change the score bOthelloGrid
    
    If NumberOfBlack + NumberOfWhite = 144 Then    'Is grid fill
        Call message("Game Over", True, False, True, False, False)
        If NumberOfBlack > NumberOfWhite Then 'is there more of computers than users
            If PlayerCounter = 1 Then
                Call message("Computer Wins", True, False, True, False, False)
            Else
                Call message("Player Wins", True, False, True, False, False)
            End If
        ElseIf NumberOfBlack < NumberOfWhite Then 'is there more of users than computers
            If PlayerCounter = 1 Then
                Call message("Player Wins", True, False, True, False, False)
            Else
                Call message("Computer Wins", True, False, True, False, False)
            End If
        ElseIf NumberOfBlack = NumberOfWhite Then 'is there a draw
            Call message("Draw", True, False, True, False, False)
        End If
        Beep
    ElseIf NumberOfBlack = 0 Then    'Has all users been eliminated
        If PlayerCounter = 1 Then
            Call message("Player Wins", True, False, True, False, False)
        Else
            Call message("Computer Wins", True, False, True, False, False)
        End If
        Beep
    ElseIf NumberOfWhite = 0 Then    'Has all user's been eliminated
        If PlayerCounter = 1 Then
            Call message("Computer Wins", True, False, True, False, False)
        Else
            Call message("Player Wins", True, False, True, False, False)
        End If
        Beep
    End If
End Sub
