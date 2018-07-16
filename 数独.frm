VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   Caption         =   "数独小游戏"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9390
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "数独规则"
      Height          =   975
      Left            =   7200
      TabIndex        =   90
      Top             =   7080
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      Caption         =   "高级"
      Height          =   495
      Left            =   7560
      TabIndex        =   89
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "初级"
      Height          =   495
      Left            =   7560
      TabIndex        =   86
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "中级"
      Height          =   375
      Left            =   7560
      TabIndex        =   87
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "消息"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   84
      Top             =   6960
      Width           =   6855
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除数据"
      Height          =   615
      Left            =   7440
      TabIndex        =   83
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "电脑解题"
      Height          =   615
      Left            =   7440
      TabIndex        =   82
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成题目"
      Height          =   615
      Left            =   7440
      TabIndex        =   81
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   80
      Left            =   6240
      TabIndex        =   80
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   79
      Left            =   5520
      TabIndex        =   79
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   78
      Left            =   4800
      TabIndex        =   78
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   77
      Left            =   3960
      TabIndex        =   77
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   76
      Left            =   3240
      TabIndex        =   76
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   75
      Left            =   2520
      TabIndex        =   75
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   74
      Left            =   1680
      TabIndex        =   74
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   73
      Left            =   960
      TabIndex        =   73
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   72
      Left            =   240
      TabIndex        =   72
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   71
      Left            =   6240
      TabIndex        =   71
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   70
      Left            =   5520
      TabIndex        =   70
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   69
      Left            =   4800
      TabIndex        =   69
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   68
      Left            =   3960
      TabIndex        =   68
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   67
      Left            =   3240
      TabIndex        =   67
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   66
      Left            =   2520
      TabIndex        =   66
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   65
      Left            =   1680
      TabIndex        =   65
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   64
      Left            =   960
      TabIndex        =   64
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   63
      Left            =   240
      TabIndex        =   63
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   62
      Left            =   6240
      TabIndex        =   62
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   61
      Left            =   5520
      TabIndex        =   61
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   60
      Left            =   4800
      TabIndex        =   60
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   59
      Left            =   3960
      TabIndex        =   59
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   58
      Left            =   3240
      TabIndex        =   58
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   57
      Left            =   2520
      TabIndex        =   57
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   56
      Left            =   1680
      TabIndex        =   56
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   55
      Left            =   960
      TabIndex        =   55
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   54
      Left            =   240
      TabIndex        =   54
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   53
      Left            =   6240
      TabIndex        =   53
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   52
      Left            =   5520
      TabIndex        =   52
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   51
      Left            =   4800
      TabIndex        =   51
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   50
      Left            =   3960
      TabIndex        =   50
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   49
      Left            =   3240
      TabIndex        =   49
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   48
      Left            =   2520
      TabIndex        =   48
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   47
      Left            =   1680
      TabIndex        =   47
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   46
      Left            =   960
      TabIndex        =   46
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   45
      Left            =   240
      TabIndex        =   45
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   44
      Left            =   6240
      TabIndex        =   44
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   43
      Left            =   5520
      TabIndex        =   43
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   42
      Left            =   4800
      TabIndex        =   42
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   41
      Left            =   3960
      TabIndex        =   41
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   40
      Left            =   3240
      TabIndex        =   40
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   39
      Left            =   2520
      TabIndex        =   39
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   38
      Left            =   1680
      TabIndex        =   38
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   37
      Left            =   960
      TabIndex        =   37
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   36
      Left            =   240
      TabIndex        =   36
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   35
      Left            =   6240
      TabIndex        =   35
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   5520
      TabIndex        =   34
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   4800
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   3960
      TabIndex        =   32
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   3240
      TabIndex        =   31
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   2520
      TabIndex        =   30
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   1680
      TabIndex        =   29
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   960
      TabIndex        =   28
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   240
      TabIndex        =   27
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   6240
      TabIndex        =   26
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   5520
      TabIndex        =   25
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   4800
      TabIndex        =   24
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   3960
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   3240
      TabIndex        =   22
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   2520
      TabIndex        =   21
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   1680
      TabIndex        =   20
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   5520
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   4800
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   3960
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   3240
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   1680
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "难度"
      Height          =   2055
      Left            =   7320
      TabIndex        =   88
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line7 
      X1              =   7080
      X2              =   9360
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line6 
      X1              =   7080
      X2              =   9360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   8160
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   6855
      Left            =   120
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim number_r(0 To 8, 0 To 8), number_9(0 To 8, 0 To 8), number_m(0 To 8, 0 To 8, 0 To 8)
Dim error_number
Private Sub Clear()
For Index = 0 To 80
    Text1(Index).Text = ""
    Text1(Index).Enabled = True
    Text1(Index).BackColor = &H80000005
    Text1(Index).ForeColor = vbBlack
    Text2.Text = ""
Next Index
End Sub



Private Sub Command1_Click()
    Clear
    If Option1.Value = True Then
    hard = Int(Rnd * 21) + 20
    ElseIf Option2.Value = True Then
    hard = Int(Rnd * 13) + 40
    ElseIf Option3.Value = True Then
    hard = Int(Rnd * 11) + 50
    Else
    a = MsgBox("请选择难度！", vbOKOnly, "提示")
    If a = vbOK Then
        Exit Sub
    End If
    End If
    
Text2.Text = "正在生成数独题目，请稍后。。。"

'初始化变量
1 For row = 0 To 8
    For Column = 0 To 8
        number_r(row, Column) = 0
        xx = Int(Column / 3) + Int(row / 3) * 3
        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
        number_9(xx, yy) = 0
        Text1(row * 9 + Column).Text = ""
        Text1(row * 9 + Column).Enabled = True
        For i = 0 To 8
            number_m(row, Column, i) = 0
        Next i
    Next Column
Next row
Dim enanble(0 To 8, 0 To 8)
'随机生成三个对角线的小矩阵
        xx_s = 0
       For a = 0 To 2
            For Column = 0 To 8
1001        number_will = Int(Rnd * 9) + 1
            For i = 0 To 8
                If number_will = number_9(xx_s, i) Then
                    GoTo 1001
               End If
            Next i
            number_9(xx_s, Column) = number_will
            number_r((Int(xx_s / 3) * 3 + Int(Column / 3)), ((xx_s Mod 3) * 3 + (Column Mod 3))) = number_9(xx_s, Column)
            enanble((Int(xx_s / 3) * 3 + Int(Column / 3)), ((xx_s Mod 3) * 3 + (Column Mod 3))) = 1
        Next Column
            xx_s = xx_s + 4
        Next a
        


'解数独
n = 2 '标记回滚时，column是往前还是往后
        For row = 0 To 8
            For Column = 0 To 8
1000            If enanble(row, Column) = 1 Then
                        If (n Mod 2) = 0 Then
                            GoTo 1050
                        End If
                        If (n Mod 2) = 1 Then
                            GoTo 1030
                        End If
                End If

               For one_nine = 1 To 9
                    For Columns = 0 To 8
                        If one_nine = number_r(row, Columns) Then
                            GoTo 1010
                        End If
                    Next Columns

                    For rows = 0 To 8
                        If one_nine = number_r(rows, Column) Then
                           GoTo 1010
                        End If
                    Next rows

                    xx = Int(Column / 3) + Int(row / 3) * 3
                    yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
                    For yys = 0 To 8
                        If one_nine = number_9(xx, yys) Then
                            GoTo 1010
                        End If
                    Next yys

                    For i = 0 To 8
                        If number_m(row, Column, i) = one_nine Then
                            GoTo 1010
                        End If
                    Next i

                    number_r(row, Column) = one_nine
                    number_9(xx, yy) = number_r(row, Column)
                    Text1(row * 9 + Column).Text = number_r(row, Column)
                    n = 2
                    For i = 0 To 8
                        If number_m(row, Column, i) = 0 Then
                            number_m(row, Column, i) = one_nine
                            GoTo 1050
                        End If
                    Next i

1010                If one_nine = 9 Then
                        If Column = 0 And row = 0 Then
                            GoTo 1
                            Exit Sub
                        End If
                        n = 3
                        For i = 0 To 8
                            number_m(row, Column, i) = 0
                        Next i
                        number_r(row, Column) = ""
                        xx = Int(Column / 3) + Int(row / 3) * 3
                        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
                        number_9(xx, yy) = number_r(row, Column)
1030                    If Column = 0 Then
                            Column = 8
                            row = row - 1
                            GoTo 1000
                        End If
                        Column = Column - 1
                        GoTo 1000
                    End If
1020            Next one_nine
1050     Next Column
    Next row

        
    For hard_r = 1 To hard
5000        row = Int(Rnd * 9)
        Column = Int(Rnd * 9)
        row_pick = 0
        column_pick = 0
        nine_pick = 0
        xx = Int(Column / 3) + Int(row / 3) * 3
        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
        For i = 0 To 8
            If number_r(i, Column) <> "" Then
                row_pick = row_pick + 1
            End If
            If number_r(row, i) <> "" Then
                column_pick = column_pick + 1
            End If
            If number_9(xx, i) <> "" Then
                nine_pick = nine_pick + 1
            End If
        Next i
        If nine_pick = 1 Or row_pick = 1 Or column_pick = 1 Then
            GoTo 5000
        End If
        If number_r(row, Column) <> "" Then
            number_r(row, Column) = ""
        Else
            GoTo 5000
        End If
    Next hard_r

    For row = 0 To 8
        For Column = 0 To 8
            Text1(row * 9 + Column).Text = number_r(row, Column)
            If number_r(row, Column) <> "" Then
            Text1(row * 9 + Column).Enabled = False
            End If
        Next Column, row

    Text2.Text = "题目生成成功"





End Sub

Private Sub Command2_Click()
'暴力解数独



'处理数组，导入待解数独
For row = 0 To 8
    For Column = 0 To 8
        number_r(row, Column) = ""
        number_r(row, Column) = Int(Val(Text1(row * 9 + Column).Text))
        xx = Int(Column / 3) + Int(row / 3) * 3
        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
        number_9(xx, yy) = number_r(row, Column)
        For i = 0 To 8
            number_m(row, Column, i) = 0
        Next i
    Next Column
Next row
Dim enanble(0 To 8, 0 To 8)
For Index = 0 To 80
    If Text1(Index).Text <> "" Then
        enanble(Int(Index / 9), (Index + 9) Mod 9) = 1
    End If
Next Index


'解数独
n = 2 '标记回滚时，column是往前还是往后
        For row = 0 To 8
            For Column = 0 To 8
3000            If enanble(row, Column) = 1 Then
                        If (n Mod 2) = 0 Then
                            GoTo 3050
                        End If
                        If (n Mod 2) = 1 Then
                            GoTo 3030
                        End If
                End If

               For one_nine = 1 To 9
                    For Columns = 0 To 8
                        If one_nine = number_r(row, Columns) Then
                            GoTo 3010
                        End If
                    Next Columns

                    For rows = 0 To 8
                        If one_nine = number_r(rows, Column) Then
                           GoTo 3010
                        End If
                    Next rows

                    xx = Int(Column / 3) + Int(row / 3) * 3
                    yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
                    For yys = 0 To 8
                        If one_nine = number_9(xx, yys) Then
                            GoTo 3010
                        End If
                    Next yys

                    For i = 0 To 8
                        If number_m(row, Column, i) = one_nine Then
                            GoTo 3010
                        End If
                    Next i

                    number_r(row, Column) = one_nine
                    number_9(xx, yy) = number_r(row, Column)
                    Text1(row * 9 + Column).Text = number_r(row, Column)
                    n = 2
                    For i = 0 To 8
                        If number_m(row, Column, i) = 0 Then
                            number_m(row, Column, i) = one_nine
                            GoTo 3050
                        End If
                    Next i

3010                If one_nine = 9 Then
                        If Column = 0 And row = 0 Then
                            y = MsgBox("此数独无解", vbOKOnly, "提示")
                            Text2.Text = "此数独无解"
                            Exit Sub
                        End If
                        n = 3
                        For i = 0 To 8
                            number_m(row, Column, i) = 0
                        Next i
                        number_r(row, Column) = ""
                        xx = Int(Column / 3) + Int(row / 3) * 3
                        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
                        number_9(xx, yy) = number_r(row, Column)
3030                    If Column = 0 Then
                            Column = 8
                            row = row - 1
                            GoTo 3000
                        End If
                        Column = Column - 1
                        GoTo 3000
                    End If
3020            Next one_nine
3050     Next Column
    Next row
    '输出结果
    For row = 0 To 8
        For Column = 0 To 8
            If enanble(row, Column) = 0 Then
                Text1(row * 9 + Column).Text = number_r(row, Column)
                Text1(row * 9 + Column).BackColor = &HFF8080

            End If
        Next Column
    Next row
    
    Text2.Text = "解题完毕"
End Sub

Private Sub Command3_Click()
Clear
End Sub


Private Sub Command4_Click()
Dialog.Show
Form1.Hide

End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).MaxLength = 1
Text1(Index).BackColor = &HC0FFC0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
2  Select Case KeyCode
    Case 37 '左
        Index = Index - 1
    Case 38 '上
        Index = Index - 9
    Case 39 '右
        Index = Index + 1
    Case 40 '下
        Index = Index + 9
End Select
If Index > 80 Then
    Index = 80
End If
If Index < 0 Then
    Index = 0
End If
If Text1(Index).Enabled = False Then
If Index = 0 Or Index = 80 Then
Exit Sub
End If
 GoTo 2
 End If
Text1(Index).SetFocus
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If
If KeyAscii < 37 Or (KeyAscii > 41 And KeyAscii < 49) Or KeyAscii > 57 Then
KeyAscii = 0
Exit Sub
End If
4 If Index < 80 Then
    Index = Index + 1
    If Text1(Index).Enabled = False Then
     GoTo 4
    End If
    Text1(Index).SetFocus
End If

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'检查错误

For row = 0 To 8
    For Column = 0 To 8
        number_r(row, Column) = 0
        number_r(row, Column) = Int(Val(Text1(row * 9 + Column).Text))
        xx = Int(Column / 3) + Int(row / 3) * 3
        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
        number_9(xx, yy) = number_r(row, Column)
    Next Column
Next row

error_number = 0
For row = 0 To 8
    For Column = 0 To 8
        If number_r(row, Column) = 0 Then
            GoTo 4000
        End If
        For Columns = 0 To 8
            If number_r(row, Column) = number_r(row, Columns) And Columns <> Column Then
                error_number = 1
                Text1(row * 9 + Column).ForeColor = vbRed
                Text1(row * 9 + Columns).ForeColor = vbRed
            End If
        Next Columns
                
        For rows = 0 To 8
            If number_r(row, Column) = number_r(rows, Column) And rows <> row Then
                error_number = 1
                Text1(row * 9 + Column).ForeColor = vbRed
                Text1(rows * 9 + Column).ForeColor = vbRed
            End If
        Next rows
                
        xx = Int(Column / 3) + Int(row / 3) * 3
        yy = (Column - Int(Column / 3) * 3) + (row - Int(row / 3) * 3) * 3
        row_xx = (Int(xx / 3) * 3 + Int(yy / 3))
        Column_yy = ((xx Mod 3) * 3 + (yy Mod 3))

        For yys = 0 To 8
            If number_9(xx, yys) = number_r(row, Column) And yy <> yys Then
                error_number = 1
                Text1(Index).ForeColor = vbRed
                Text1((Int(xx / 3) * 3 + Int(yys / 3)) * 9 + ((xx Mod 3) * 3 + (yys Mod 3))).ForeColor = vbRed
            End If
        Next yys
4000    Next Column
    
Next row

    If error_number = 1 Then
        Text2.Text = "输入的字有重复"
    Else
        Text2.Text = ""
        For i = 0 To 80
            Text1(i).ForeColor = vbBlack
        Next i
    End If
    
 Number = 0
    For row = 0 To 8
        For Column = 0 To 8
            If number_r(row, Column) > 0 Then
                Number = Number + 1
            End If
    Next Column, row
     If Number = 81 And error_number = 0 Then
        x = MsgBox("数独填写完成！", vbOKOnly, "恭喜")
        For Index = 0 To 80
            Text1(Index).Enabled = False
        Next Index
    End If
        
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &HFFFFFF
End Sub
