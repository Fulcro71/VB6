VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Termina"
      Height          =   495
      Left            =   7920
      TabIndex        =   82
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Frame Frame8 
      Caption         =   "Meccanismo di rottura"
      Height          =   4935
      Left            =   5640
      TabIndex        =   67
      Top             =   240
      Width           =   3975
      Begin VB.Frame Frame9 
         Caption         =   "Sollecitazioni"
         Height          =   1335
         Left            =   1800
         TabIndex        =   93
         Top             =   1200
         Width           =   1935
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text24 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text23 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "N/mm^2"
            Height          =   195
            Left            =   1200
            TabIndex        =   105
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "N/mm^2"
            Height          =   195
            Left            =   1200
            TabIndex        =   104
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "N/mm^2"
            Height          =   195
            Left            =   1200
            TabIndex        =   103
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "s'"
            Height          =   195
            Left            =   240
            TabIndex        =   99
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   720
            Width           =   75
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "c"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "s    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   102
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "s    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   101
            Top             =   600
            Width           =   330
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "s    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Deformazioni"
         Height          =   1335
         Left            =   120
         TabIndex        =   83
         Top             =   1200
         Width           =   1575
         Begin VB.TextBox Text22 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "s'"
            Height          =   195
            Left            =   480
            TabIndex        =   89
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   480
            TabIndex        =   88
            Top             =   720
            Width           =   75
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "e    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   92
            Top             =   960
            Width           =   315
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "e    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   91
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "c"
            Height          =   195
            Left            =   480
            TabIndex        =   87
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "e    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   90
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   720
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   3720
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Momento Ultimo di Calcolo"
         Height          =   195
         Left            =   720
         TabIndex        =   81
         Top             =   4080
         Width           =   1875
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "KNm"
         Height          =   195
         Left            =   2160
         TabIndex        =   80
         Top             =   4440
         Width           =   345
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "Rd"
         Height          =   195
         Left            =   600
         TabIndex        =   78
         Top             =   4560
         Width           =   210
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "M     :"
         Height          =   195
         Left            =   480
         TabIndex        =   77
         Top             =   4440
         Width           =   405
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   76
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "x :"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   75
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Asse neutro:"
         Height          =   195
         Left            =   600
         TabIndex        =   74
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   2400
         TabIndex        =   73
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "(yn/d)"
         Height          =   195
         Left            =   2400
         TabIndex        =   72
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label56 
         Height          =   495
         Left            =   480
         TabIndex        =   69
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Campo di collasso :"
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3720
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Campo di rottura"
      Height          =   495
      Left            =   4200
      TabIndex        =   48
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcola Parametri"
      Height          =   495
      Left            =   2400
      TabIndex        =   47
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Caratteristiche dei materiali"
      Height          =   4935
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Width           =   3015
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Variabili di calcolo"
         Height          =   1695
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   2775
         Begin VB.TextBox Text29 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label66 
            Caption         =   "fcd"
            Height          =   255
            Left            =   1680
            TabIndex        =   108
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   107
            Top             =   200
            Width           =   120
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "s'"
            Height          =   195
            Left            =   1800
            TabIndex        =   57
            Top             =   1440
            Width           =   105
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "yd"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   1440
            Width           =   165
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "fcd:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "fyd:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "ftd:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   225
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "e     :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "d:"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   61
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "r    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1680
            TabIndex        =   59
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   1800
            TabIndex        =   58
            Top             =   1080
            Width           =   75
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "r    :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1680
            TabIndex        =   60
            Top             =   960
            Width           =   315
         End
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1320
         TabIndex        =   45
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1440
         TabIndex        =   41
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         TabIndex        =   37
         Text            =   "Combo3"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Text            =   "30"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "N/mm^2"
         Height          =   195
         Left            =   2160
         TabIndex        =   46
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Es:"
         Height          =   195
         Left            =   960
         TabIndex        =   44
         Top             =   2640
         Width           =   225
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "N/mm^2"
         Height          =   195
         Left            =   2160
         TabIndex        =   43
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "N/mm^2"
         Height          =   195
         Left            =   2160
         TabIndex        =   42
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "ftk:"
         Height          =   195
         Left            =   1080
         TabIndex        =   39
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "fyk:"
         Height          =   195
         Left            =   1080
         TabIndex        =   38
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Acciaio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   36
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "N/mm^2"
         Height          =   195
         Left            =   2160
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "N/mm^2"
         Height          =   195
         Left            =   2160
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Rck:"
         Height          =   195
         Left            =   960
         TabIndex        =   32
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "fck:"
         Height          =   195
         Left            =   1080
         TabIndex        =   30
         Top             =   600
         Width           =   270
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Calcestruzzo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Caratteristiche della sezione"
      Height          =   1575
      Left            =   3240
      TabIndex        =   18
      Top             =   240
      Width           =   2295
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "4"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Text            =   "31"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Text            =   "22"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   1920
         TabIndex        =   27
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Copriferro:"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "cm"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Altezza utile:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Base:"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Armatura tesa"
      Height          =   1575
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Text            =   "4"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "cm^2"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Area:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Diametro:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° barre:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Armatura compressa"
      Height          =   1575
      Left            =   3240
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "2"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "cm^2"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Area:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Diametro:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° barre:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Text1.Text > "0" And Text1.Text < "9" And Text1.Text <> "" Then
            Afc = Int(Text1.Text) * (Int(Combo1.List(Combo1.ListIndex)) / 20) ^ 2 * 3.141592654
            Text2.Text = Format$(Afc, "##.##")
            Else
            Text2.Text = ""
            
            End If
            Call Resetta
End Sub



Private Sub Combo2_Click()
If Text3.Text > "0" And Text3.Text < "9" And Text3.Text <> "" Then
            Aft = Int(Text3.Text) * (Int(Combo2.List(Combo2.ListIndex)) / 20) ^ 2 * 3.141592654
            Text4.Text = Format$(Aft, "##.##")
            Else
            Text4.Text = ""
            
            End If
            Call Resetta
End Sub

Private Sub Combo3_Click()
Fe = Combo3.ListIndex
If Fe = 0 Then
            fyk = 375
            ftk = 450
            End If
If Fe = 1 Then
            fyk = 430
            ftk = 540
            End If
Text10.Text = fyk
Text11.Text = ftk
Call Resetta
End Sub

Private Sub Command1_Click()
Call Assign
End Sub

Private Sub Command2_Click()
C2 = Campo2
If C2 > 0 Then
        DefStatus (C2)
        Label56.Caption = "- Collasso dell'acciaio teso." + Chr(10) + Chr(13) + "- Calcestruzzo non al collasso"
        Label57.Caption = "Campo 2"
        Call Mrd
        Exit Sub
        End If
C3 = Campo3
If C3 > 0 Then
        DefStatus (C3)
        Label56.Caption = "- Collasso del calcestruzzo." + Chr(10) + Chr(13) + "- Acciaio teso in campo plastico."
        Label57.Caption = "Campo 3"
        Call Mrd
        Exit Sub
        End If
C4 = Campo4
If C4 > 0 Then
        DefStatus (C4)
        Label56.Caption = "- Collasso del calcestruzzo." + Chr(10) + Chr(13) + "- Acciaio teso in campo elastico."
        Label57.Caption = "Campo 4"
        Call Mrd
        Exit Sub
        End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Form1.Caption = "Stato Limite Ultimo per Flessione"
Command2.Enabled = False
For Diam = 6 To 30 Step 2
Combo1.AddItem Diam
Next Diam

For Diam = 6 To 30 Step 2
Combo2.AddItem Diam
Next Diam
Combo3.AddItem "FeB38k"
Combo3.AddItem "FeB44k"

Combo1.ListIndex = 3
Combo2.ListIndex = 3
Combo3.ListIndex = 1
Text12.Text = Eacciaio
fck = Text8.Text
    Rck = fck / 0.83
    Text9.Text = Format$(Rck, "##.##")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_Change()
If Text1.Text > "0" And Text1.Text < "9" And Text1.Text <> "" Then
            Afc = Int(Text1.Text) * (Int(Combo1.List(Combo1.ListIndex)) / 20) ^ 2 * 3.141592654
            Text2.Text = Format$(Afc, "##.##")
            Else
            Text2.Text = ""
            
            End If
            Call Resetta
End Sub

Private Sub Text3_Change()

If Text3.Text > "0" And Text3.Text < "9" And Text3.Text <> "" Then
            Aft = Int(Text3.Text) * (Int(Combo2.List(Combo2.ListIndex)) / 20) ^ 2 * 3.141592654
            Text4.Text = Format$(Aft, "##.##")
            Else
            Text4.Text = ""
            
            End If
            Call Resetta
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Call Resetta
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
    Call Resetta
End Sub

Private Sub Text7_Change()
Call Resetta
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    fck = Text8.Text
    Rck = fck / 0.83
    Text9.Text = Format$(Rck, "##.##")
    Call Resetta
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Rck = Text9.Text
    fck = Rck * 0.83
    Text8.Text = Format$(fck, "##.##")
    Call Resetta
    End If
End Sub
