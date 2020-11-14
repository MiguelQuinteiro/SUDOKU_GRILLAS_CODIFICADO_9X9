VERSION 5.00
Begin VB.Form frmCODIFICADO 
   BackColor       =   &H00800000&
   Caption         =   "SUDOKU GRILLAS CODIFICADO"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSiguiente 
      BackColor       =   &H00FF8080&
      Caption         =   "SIGUIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H00FF8080&
      Caption         =   "ANTERIOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   6240
      TabIndex        =   132
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5760
      TabIndex        =   131
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   5280
      TabIndex        =   130
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6240
      TabIndex        =   129
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5760
      TabIndex        =   128
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5280
      TabIndex        =   127
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6240
      TabIndex        =   126
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5760
      TabIndex        =   125
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoRegional 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   124
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4560
      TabIndex        =   122
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4080
      TabIndex        =   121
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3600
      TabIndex        =   120
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4560
      TabIndex        =   119
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4080
      TabIndex        =   118
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3600
      TabIndex        =   117
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4560
      TabIndex        =   116
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   115
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoVertical 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   114
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   9
      Left            =   2880
      TabIndex        =   112
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   8
      Left            =   2400
      TabIndex        =   111
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   1920
      TabIndex        =   110
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   6
      Left            =   2880
      TabIndex        =   109
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   5
      Left            =   2400
      TabIndex        =   108
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   1920
      TabIndex        =   107
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   2880
      TabIndex        =   106
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2400
      TabIndex        =   105
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigoHorizontal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1920
      TabIndex        =   104
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   81
      Left            =   11160
      TabIndex        =   103
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   10680
      TabIndex        =   102
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   79
      Left            =   10200
      TabIndex        =   101
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   78
      Left            =   9600
      TabIndex        =   100
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   9120
      TabIndex        =   99
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   76
      Left            =   8640
      TabIndex        =   98
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   75
      Left            =   8040
      TabIndex        =   97
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   7560
      TabIndex        =   96
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   73
      Left            =   7080
      TabIndex        =   95
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   72
      Left            =   11160
      TabIndex        =   94
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   71
      Left            =   10680
      TabIndex        =   93
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   70
      Left            =   10200
      TabIndex        =   92
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   69
      Left            =   9600
      TabIndex        =   91
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   68
      Left            =   9120
      TabIndex        =   90
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   67
      Left            =   8640
      TabIndex        =   89
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   66
      Left            =   8040
      TabIndex        =   88
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   65
      Left            =   7560
      TabIndex        =   87
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   64
      Left            =   7080
      TabIndex        =   86
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   63
      Left            =   11160
      TabIndex        =   85
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   62
      Left            =   10680
      TabIndex        =   84
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   61
      Left            =   10200
      TabIndex        =   83
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   60
      Left            =   9600
      TabIndex        =   82
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   59
      Left            =   9120
      TabIndex        =   81
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   58
      Left            =   8640
      TabIndex        =   80
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   57
      Left            =   8040
      TabIndex        =   79
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   56
      Left            =   7560
      TabIndex        =   78
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   55
      Left            =   7080
      TabIndex        =   77
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   54
      Left            =   11160
      TabIndex        =   76
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   10680
      TabIndex        =   75
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   52
      Left            =   10200
      TabIndex        =   74
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   51
      Left            =   9600
      TabIndex        =   73
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   9120
      TabIndex        =   72
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   49
      Left            =   8640
      TabIndex        =   71
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   48
      Left            =   8040
      TabIndex        =   70
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   7560
      TabIndex        =   69
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   46
      Left            =   7080
      TabIndex        =   68
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   45
      Left            =   11160
      TabIndex        =   67
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   44
      Left            =   10680
      TabIndex        =   66
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   43
      Left            =   10200
      TabIndex        =   65
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   42
      Left            =   9600
      TabIndex        =   64
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   41
      Left            =   9120
      TabIndex        =   63
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   40
      Left            =   8640
      TabIndex        =   62
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   39
      Left            =   8040
      TabIndex        =   61
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   38
      Left            =   7560
      TabIndex        =   60
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   37
      Left            =   7080
      TabIndex        =   59
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   36
      Left            =   11160
      TabIndex        =   58
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   35
      Left            =   10680
      TabIndex        =   57
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   34
      Left            =   10200
      TabIndex        =   56
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   33
      Left            =   9600
      TabIndex        =   55
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   32
      Left            =   9120
      TabIndex        =   54
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   31
      Left            =   8640
      TabIndex        =   53
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   30
      Left            =   8040
      TabIndex        =   52
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   29
      Left            =   7560
      TabIndex        =   51
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   28
      Left            =   7080
      TabIndex        =   50
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   27
      Left            =   11160
      TabIndex        =   49
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   10680
      TabIndex        =   48
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   25
      Left            =   10200
      TabIndex        =   47
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   24
      Left            =   9600
      TabIndex        =   46
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   9120
      TabIndex        =   45
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   22
      Left            =   8640
      TabIndex        =   44
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   21
      Left            =   8040
      TabIndex        =   43
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   7560
      TabIndex        =   42
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   19
      Left            =   7080
      TabIndex        =   41
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   18
      Left            =   11160
      TabIndex        =   40
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   17
      Left            =   10680
      TabIndex        =   39
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   16
      Left            =   10200
      TabIndex        =   38
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   15
      Left            =   9600
      TabIndex        =   37
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   14
      Left            =   9120
      TabIndex        =   36
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   13
      Left            =   8640
      TabIndex        =   35
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   12
      Left            =   8040
      TabIndex        =   34
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   11
      Left            =   7560
      TabIndex        =   33
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   10
      Left            =   7080
      TabIndex        =   32
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   9
      Left            =   11160
      TabIndex        =   31
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   8
      Left            =   10680
      TabIndex        =   30
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   7
      Left            =   10200
      TabIndex        =   29
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   28
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   5
      Left            =   9120
      TabIndex        =   27
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   4
      Left            =   8640
      TabIndex        =   26
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   8040
      TabIndex        =   25
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   2
      Left            =   7560
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCasilla 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   23
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   9
      Left            =   1200
      TabIndex        =   19
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   18
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   16
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   15
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   13
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5760
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5160
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3960
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtGrilla 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtMuestraGrilla 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdMuestraGrilla 
      BackColor       =   &H00FF8080&
      Caption         =   "MUESTRA GRILLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "DIAGRAMA EN EL SUDOKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   134
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "REGIONAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   133
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "VERTICAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3600
      TabIndex        =   123
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "HORIZONTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   113
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "NUMERO A CONSULTAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   22
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "GRILLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   1920
      Width           =   5295
   End
End
Attribute VB_Name = "frmCODIFICADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : SUDOKU GRILLAS CODIFICADO
'* CONTENIDO     : MUESTRA LOS CODIGOS Y LAS GRILLAS ASOCIADAS
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 15 DE DICIEMBRE DE 2013
'* ACTUALIZACION : 15 DE DICIEMBRE DE 2013
'****************************************************************************************
Option Explicit
Option Base 1

' DECLARACION DE TIPOS DE DATOS CREADOS POR EL USUARIO
Private Type LasGrillas
  NumeroGrilla As Long
  DigitoGrilla(1 To 9) As Integer
  CodigoGrilla(1 To 9) As Integer
  CodigoHorizontal(1 To 9) As String
  CodigoVertical(1 To 9) As String
  CodigoRegional(1 To 9) As String
  ActivaGrilla As Boolean
End Type

Private Type LasCasilla
  miContadorColorCasilla As Integer
  miDigitoColor As Integer
End Type

' DEFINICION DE VARIABLES BASADAS EN TIPOS DE DATOS CREADOS POR EL USUARIO
Dim miGrilla(1 To 46656) As LasGrillas
Dim miGrilla1(1 To 46656) As LasGrillas
Dim miGrilla2(1 To 46656) As LasGrillas
Dim miGrilla3(1 To 46656) As LasGrillas
Dim miGrilla4(1 To 46656) As LasGrillas
Dim miGrilla5(1 To 46656) As LasGrillas
Dim miGrilla6(1 To 46656) As LasGrillas
Dim miGrilla7(1 To 46656) As LasGrillas
Dim miGrilla8(1 To 46656) As LasGrillas
Dim miGrilla9(1 To 46656) As LasGrillas
Dim miCasilla(1 To 81) As LasCasilla

' DEFINICION DE VARIABLES
Dim miLineInput As String
Dim miLineOutput As String
Dim miNumeroGrilla As Long
Dim miCodigoGrilla As Long
Dim miConsultaGrilla As Long



Private Sub cmdAnterior_Click()
  If miConsultaGrilla > 1 Then
    miConsultaGrilla = miConsultaGrilla - 1
    txtMuestraGrilla.Text = miConsultaGrilla
    Call cmdMuestraGrilla_Click
  Else
    MsgBox (" YA ESTAS EN LA PRIMERA GRILLA ")
  End If

End Sub

Private Sub cmdSiguiente_Click()
  If miConsultaGrilla < 46656 Then
    miConsultaGrilla = miConsultaGrilla + 1
    txtMuestraGrilla.Text = miConsultaGrilla
    Call cmdMuestraGrilla_Click
  Else
    MsgBox (" YA ESTAS EN LA ÚLTIMA GRILLA ")
  End If
End Sub

' AL MOMENTO DE CARGAR EL FORMULARIO
Private Sub Form_Load()
  Dim i As Integer

  ' ABRE EL ARCHIVO CON LAS 46.656 GRILLAS ORIGINALES
  Open "LasGrillas.txt" For Input As #10
  Do Until EOF(10)
    Line Input #10, miLineInput
    miNumeroGrilla = Val(Mid(miLineInput, 35, 5))

    ' CARGA EL NUMERO DE LA GRILLA
    miGrilla(miNumeroGrilla).NumeroGrilla = Val(Mid(miLineInput, 35, 5))
    ' CARGA LOS VALORES DE LOS DÍGITOS QUE COMPONEN LA GRILLA
    miGrilla(miNumeroGrilla).DigitoGrilla(1) = Val(Mid(miLineInput, 2, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(2) = Val(Mid(miLineInput, 5, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(3) = Val(Mid(miLineInput, 8, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(4) = Val(Mid(miLineInput, 11, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(5) = Val(Mid(miLineInput, 14, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(6) = Val(Mid(miLineInput, 17, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(7) = Val(Mid(miLineInput, 20, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(8) = Val(Mid(miLineInput, 23, 2))
    miGrilla(miNumeroGrilla).DigitoGrilla(9) = Val(Mid(miLineInput, 26, 2))
    ' CARGA EL CODIGO DE LA GRILLA
    miGrilla(miNumeroGrilla).CodigoGrilla(1) = Val(Mid(miLineInput, 42, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(2) = Val(Mid(miLineInput, 43, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(3) = Val(Mid(miLineInput, 44, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(4) = Val(Mid(miLineInput, 45, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(5) = Val(Mid(miLineInput, 46, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(6) = Val(Mid(miLineInput, 47, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(7) = Val(Mid(miLineInput, 48, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(8) = Val(Mid(miLineInput, 49, 1))
    miGrilla(miNumeroGrilla).CodigoGrilla(9) = Val(Mid(miLineInput, 50, 1))
    ' CARGA EL CODIGO HORIZONAL DE LA GRILLA
    miGrilla(miNumeroGrilla).CodigoHorizontal(1) = Mid(miLineInput, 53, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(2) = Mid(miLineInput, 54, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(3) = Mid(miLineInput, 55, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(4) = Mid(miLineInput, 56, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(5) = Mid(miLineInput, 57, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(6) = Mid(miLineInput, 58, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(7) = Mid(miLineInput, 59, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(8) = Mid(miLineInput, 60, 1)
    miGrilla(miNumeroGrilla).CodigoHorizontal(9) = Mid(miLineInput, 61, 1)
    ' CARGA EL CODIGO VERTICAL DE LA GRILLA
    miGrilla(miNumeroGrilla).CodigoVertical(1) = Mid(miLineInput, 64, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(2) = Mid(miLineInput, 65, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(3) = Mid(miLineInput, 66, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(4) = Mid(miLineInput, 67, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(5) = Mid(miLineInput, 68, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(6) = Mid(miLineInput, 69, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(7) = Mid(miLineInput, 70, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(8) = Mid(miLineInput, 71, 1)
    miGrilla(miNumeroGrilla).CodigoVertical(9) = Mid(miLineInput, 72, 1)
    ' CARGA EL CODIGO REGIONAL DE LA GRILLA
    miGrilla(miNumeroGrilla).CodigoRegional(1) = Mid(miLineInput, 75, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(2) = Mid(miLineInput, 76, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(3) = Mid(miLineInput, 77, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(4) = Mid(miLineInput, 78, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(5) = Mid(miLineInput, 79, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(6) = Mid(miLineInput, 80, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(7) = Mid(miLineInput, 81, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(8) = Mid(miLineInput, 82, 1)
    miGrilla(miNumeroGrilla).CodigoRegional(9) = Mid(miLineInput, 83, 1)
    ' ACTIVA LA GRILLA PARA LA SOLUCION DEL PROBLEMA
    miGrilla(miNumeroGrilla).ActivaGrilla = True
  Loop
  Close #10
End Sub


Private Sub cmdMuestraGrilla_Click()
' DEFINICION DE VARIABLES
  Dim i As Integer

  ' EXTRAE EL VALOR DE LA GRILLA A CONSULTAR
  miConsultaGrilla = Val(txtMuestraGrilla)

  ' MUESTRA LA CONSULTA SI EL VALOR ESTA ENTRE 1 Y 46656
  If miConsultaGrilla > 0 And miConsultaGrilla < 46657 Then
    ' LIMPIA LOS DATOS
    For i = 1 To 9
      txtGrilla(i) = ""
      txtCodigo(i) = ""
    Next i

    For i = 1 To 81
      txtCasilla(i) = ""
    Next i

    ' MUESTRA LA GRILLA EN NUMEROS
    txtGrilla(1) = miGrilla(miConsultaGrilla).DigitoGrilla(1)
    txtGrilla(2) = miGrilla(miConsultaGrilla).DigitoGrilla(2)
    txtGrilla(3) = miGrilla(miConsultaGrilla).DigitoGrilla(3)
    txtGrilla(4) = miGrilla(miConsultaGrilla).DigitoGrilla(4)
    txtGrilla(5) = miGrilla(miConsultaGrilla).DigitoGrilla(5)
    txtGrilla(6) = miGrilla(miConsultaGrilla).DigitoGrilla(6)
    txtGrilla(7) = miGrilla(miConsultaGrilla).DigitoGrilla(7)
    txtGrilla(8) = miGrilla(miConsultaGrilla).DigitoGrilla(8)
    txtGrilla(9) = miGrilla(miConsultaGrilla).DigitoGrilla(9)

    ' MUESTRA LA GRILLA EN DIAGRAMA SUDOKU
    txtCasilla(Val(txtGrilla(1))) = "@"
    txtCasilla(Val(txtGrilla(2))) = "@"
    txtCasilla(Val(txtGrilla(3))) = "@"
    txtCasilla(Val(txtGrilla(4))) = "@"
    txtCasilla(Val(txtGrilla(5))) = "@"
    txtCasilla(Val(txtGrilla(6))) = "@"
    txtCasilla(Val(txtGrilla(7))) = "@"
    txtCasilla(Val(txtGrilla(8))) = "@"
    txtCasilla(Val(txtGrilla(9))) = "@"

    ' MUESTRA EL CODIGO
    txtCodigo(1) = miGrilla(miConsultaGrilla).CodigoGrilla(1)
    txtCodigo(2) = miGrilla(miConsultaGrilla).CodigoGrilla(2)
    txtCodigo(3) = miGrilla(miConsultaGrilla).CodigoGrilla(3)
    txtCodigo(4) = miGrilla(miConsultaGrilla).CodigoGrilla(4)
    txtCodigo(5) = miGrilla(miConsultaGrilla).CodigoGrilla(5)
    txtCodigo(6) = miGrilla(miConsultaGrilla).CodigoGrilla(6)
    txtCodigo(7) = miGrilla(miConsultaGrilla).CodigoGrilla(7)
    txtCodigo(8) = miGrilla(miConsultaGrilla).CodigoGrilla(8)
    txtCodigo(9) = miGrilla(miConsultaGrilla).CodigoGrilla(9)

    ' MUESTRA EL CODIGO HORIZONTAL
    txtCodigoHorizontal(1) = miGrilla(miConsultaGrilla).CodigoHorizontal(1)
    txtCodigoHorizontal(2) = miGrilla(miConsultaGrilla).CodigoHorizontal(2)
    txtCodigoHorizontal(3) = miGrilla(miConsultaGrilla).CodigoHorizontal(3)
    txtCodigoHorizontal(4) = miGrilla(miConsultaGrilla).CodigoHorizontal(4)
    txtCodigoHorizontal(5) = miGrilla(miConsultaGrilla).CodigoHorizontal(5)
    txtCodigoHorizontal(6) = miGrilla(miConsultaGrilla).CodigoHorizontal(6)
    txtCodigoHorizontal(7) = miGrilla(miConsultaGrilla).CodigoHorizontal(7)
    txtCodigoHorizontal(8) = miGrilla(miConsultaGrilla).CodigoHorizontal(8)
    txtCodigoHorizontal(9) = miGrilla(miConsultaGrilla).CodigoHorizontal(9)

    ' MUESTRA EL CODIGO VERTICAL
    txtCodigoVertical(1) = miGrilla(miConsultaGrilla).CodigoVertical(1)
    txtCodigoVertical(2) = miGrilla(miConsultaGrilla).CodigoVertical(2)
    txtCodigoVertical(3) = miGrilla(miConsultaGrilla).CodigoVertical(3)
    txtCodigoVertical(4) = miGrilla(miConsultaGrilla).CodigoVertical(4)
    txtCodigoVertical(5) = miGrilla(miConsultaGrilla).CodigoVertical(5)
    txtCodigoVertical(6) = miGrilla(miConsultaGrilla).CodigoVertical(6)
    txtCodigoVertical(7) = miGrilla(miConsultaGrilla).CodigoVertical(7)
    txtCodigoVertical(8) = miGrilla(miConsultaGrilla).CodigoVertical(8)
    txtCodigoVertical(9) = miGrilla(miConsultaGrilla).CodigoVertical(9)

    ' MUESTRA EL CODIGO REGIONAL
    txtCodigoRegional(1) = miGrilla(miConsultaGrilla).CodigoRegional(1)
    txtCodigoRegional(2) = miGrilla(miConsultaGrilla).CodigoRegional(2)
    txtCodigoRegional(3) = miGrilla(miConsultaGrilla).CodigoRegional(3)
    txtCodigoRegional(4) = miGrilla(miConsultaGrilla).CodigoRegional(4)
    txtCodigoRegional(5) = miGrilla(miConsultaGrilla).CodigoRegional(5)
    txtCodigoRegional(6) = miGrilla(miConsultaGrilla).CodigoRegional(6)
    txtCodigoRegional(7) = miGrilla(miConsultaGrilla).CodigoRegional(7)
    txtCodigoRegional(8) = miGrilla(miConsultaGrilla).CodigoRegional(8)
    txtCodigoRegional(9) = miGrilla(miConsultaGrilla).CodigoRegional(9)
  End If
End Sub
