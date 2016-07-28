VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Register 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Register"
   ClientHeight    =   10335
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   18255
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   18255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmManual 
      BackColor       =   &H008080FF&
      Caption         =   "Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   12480
      TabIndex        =   155
      Top             =   3120
      Width           =   3135
      Begin VB.CommandButton cmdExtraClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   158
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdExtraOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   157
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtExtra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   840
         TabIndex        =   156
         Top             =   315
         Width           =   2175
      End
      Begin VB.Label lblExtra 
         BackColor       =   &H008080FF&
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   159
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frmGiftCertificate 
      BackColor       =   &H008080FF&
      Caption         =   "GiftCertificate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   9360
      TabIndex        =   52
      Top             =   7440
      Width           =   3135
      Begin VB.CommandButton cmdSellGC 
         Caption         =   "Sell Gift Certificate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frmDiscount 
      BackColor       =   &H008080FF&
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   12480
      TabIndex        =   83
      Top             =   1320
      Width           =   3135
      Begin VB.CommandButton cmdDisCancel 
         Caption         =   "Cancel"
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
         Left            =   1440
         TabIndex        =   84
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox comDiscount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   86
         Text            =   "Select Right Discount"
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdDisOK 
         Caption         =   "OK"
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
         Left            =   1440
         TabIndex        =   85
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   975
         Left            =   120
         TabIndex        =   87
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame frmManager 
      BackColor       =   &H008080FF&
      Caption         =   "Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   15600
      TabIndex        =   88
      Top             =   1320
      Width           =   2655
      Begin VB.CommandButton cmdMainWindow 
         Cancel          =   -1  'True
         Caption         =   "Main Window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   136
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmdReview 
         Caption         =   "Review"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   89
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdTicket 
         Caption         =   "Ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   150
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdClosing 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Closing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00C0E0FF&
         TabIndex        =   90
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame frmTransactionDay 
      BackColor       =   &H008080FF&
      Height          =   1335
      Left            =   12480
      TabIndex        =   91
      Top             =   0
      Width           =   5775
      Begin VB.Timer tmrCurrentTime 
         Interval        =   100
         Left            =   1200
         Top             =   480
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   120
         TabIndex        =   93
         Top             =   840
         Width           =   5535
      End
   End
   Begin VB.Frame frmProduct 
      BackColor       =   &H008080FF&
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   9360
      TabIndex        =   51
      Top             =   8640
      Width           =   3135
      Begin VB.CommandButton cmdProductOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   126
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox comProductList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   125
         Text            =   "Select Right Product"
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frmThreading 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Threading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   0
      TabIndex        =   14
      Top             =   6120
      Width           =   3135
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   118
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   117
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdThreading 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmSpecial 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Special"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3255
      Left            =   9360
      TabIndex        =   110
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   152
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   151
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   114
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   113
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   112
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSpecial 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmTips 
      BackColor       =   &H00C0C0FF&
      Caption         =   "TipsWraps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5175
      Left            =   3120
      TabIndex        =   41
      Top             =   5160
      Width           =   3135
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   9
         Left            =   1560
         TabIndex        =   109
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   43
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   45
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   47
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTips 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmManicure 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Manicure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   11
         Left            =   1560
         TabIndex        =   124
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   10
         Left            =   120
         TabIndex        =   123
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   9
         Left            =   1560
         TabIndex        =   116
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   8
         Left            =   120
         TabIndex        =   115
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   106
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   2
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdManicure 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Tag             =   "10.8"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmReceipt 
      BackColor       =   &H00404080&
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   12480
      TabIndex        =   67
      Top             =   4680
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid msgReceipt 
         Height          =   4335
         Left            =   120
         TabIndex        =   135
         Top             =   360
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   100
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame frmTender 
         BackColor       =   &H00404080&
         Caption         =   "Tender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   4455
         Left            =   3480
         TabIndex        =   70
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   72
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdCash 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   73
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdCreditCard 
            Caption         =   "Credit Card"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   71
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmdRedemGC 
            Caption         =   "Redem Gift Certificate"
            Height          =   615
            Left            =   120
            TabIndex        =   74
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmd10 
            Caption         =   "$10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   76
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmd20 
            BackColor       =   &H00808080&
            Caption         =   "$20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmd30 
            Caption         =   "$30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   78
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmd40 
            Caption         =   "$40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmd50 
            Caption         =   "$50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   80
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd100 
            Caption         =   "$100"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtTender 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   615
            Left            =   120
            TabIndex        =   75
            Top             =   3750
            Width           =   1935
         End
         Begin VB.Label lblTenderAmount 
            BackColor       =   &H00404080&
            Caption         =   "Tendered Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   3480
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   68
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   69
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label lblBalance 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   140
         Top             =   4920
         Width           =   3135
      End
   End
   Begin VB.Frame frmPedicure 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pedicure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5175
      Left            =   3120
      TabIndex        =   8
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   9
         Left            =   1560
         TabIndex        =   122
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   8
         Left            =   120
         TabIndex        =   119
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   108
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   107
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   105
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPedicure 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmMassage 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BackMassage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3255
      Left            =   6240
      TabIndex        =   34
      Top             =   7080
      Width           =   3135
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   35
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   37
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdMassage 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmOtherServices 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OtherServices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   9360
      TabIndex        =   54
      Top             =   3240
      Width           =   3135
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   134
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   133
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   132
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   131
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   130
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   129
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   128
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdOthers 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   127
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmWaxing 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Waxing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   7095
      Left            =   6240
      TabIndex        =   21
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   13
         Left            =   1560
         TabIndex        =   121
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   12
         Left            =   120
         TabIndex        =   120
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   11
         Left            =   1560
         TabIndex        =   22
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   9
         Left            =   1560
         TabIndex        =   24
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   7
         Left            =   1560
         TabIndex        =   26
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   5
         Left            =   1560
         TabIndex        =   28
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   3
         Left            =   1560
         TabIndex        =   30
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdWaxing 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmReprint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Review"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   3360
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame frmEmpCommision 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Employee Transactions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   8655
         Left            =   8040
         TabIndex        =   95
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton cmdPrintCommision 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   148
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton cmdClearCommision 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   160
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton cmdViewEmpTran 
            Caption         =   "View Transaction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   174
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdShowCommision 
            Caption         =   "Commision/Tips"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   96
            Top             =   840
            Width           =   2055
         End
         Begin MSFlexGridLib.MSFlexGrid msgCommision 
            Height          =   6615
            Left            =   120
            TabIndex        =   142
            Top             =   1920
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   11668
            _Version        =   393216
            Rows            =   1000
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
         End
         Begin VB.ComboBox comEmpList 
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Text            =   "Select Employee"
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.CommandButton cmdCloseReprint 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   98
         Top             =   8160
         Width           =   3135
      End
      Begin MSFlexGridLib.MSFlexGrid msgSalesView 
         Height          =   8655
         Left            =   240
         TabIndex        =   141
         Top             =   360
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   15266
         _Version        =   393216
         Rows            =   10000
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Frame frmTodayTran 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Today's Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   5775
         Left            =   4680
         TabIndex        =   102
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton cmdViewAllTran 
            Caption         =   "View All Transactions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   103
            Top             =   4920
            Width           =   2895
         End
         Begin VB.CommandButton cmdPrintLastTran 
            Caption         =   "Print Last Transaction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   104
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Frame frmAdjustment 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tips Adjustment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Left            =   120
            TabIndex        =   162
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
            Begin VB.TextBox txtNewId 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   167
               Top             =   1440
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton cmdTipsCancel 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   166
               Top             =   2880
               Width           =   2655
            End
            Begin VB.CommandButton cmdTipsOK 
               Caption         =   "Update"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   165
               Top             =   2160
               Width           =   2655
            End
            Begin VB.TextBox txtTips 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   163
               Top             =   600
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.Label lblNewId 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Enter Employee ID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   168
               Top             =   1200
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.Label lblTips 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Enter Tips Amount"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   164
               Top             =   360
               Visible         =   0   'False
               Width           =   2055
            End
         End
         Begin VB.Frame frmPrintViewTran 
            BackColor       =   &H008080FF&
            Height          =   3735
            Left            =   120
            TabIndex        =   143
            Top             =   360
            Width           =   2895
            Begin VB.CommandButton cmdTipAdjustment 
               Caption         =   "Tips/ID Adjustment"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   161
               Top             =   3000
               Width           =   2655
            End
            Begin VB.CommandButton cmdVoid 
               Caption         =   "Void"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   149
               Top             =   2400
               Width           =   2655
            End
            Begin VB.CommandButton cmdViewTran 
               Caption         =   "View"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   147
               Top             =   1800
               Width           =   2655
            End
            Begin VB.CommandButton cmdPrintTran 
               Caption         =   "Print"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   146
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtTransactionNum 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   615
               Left            =   120
               TabIndex        =   145
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label lblTransactionNum 
               BackColor       =   &H008080FF&
               Caption         =   "Enter Transaction Number"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame frmUpdateIdTips 
            BackColor       =   &H00C0FFFF&
            Height          =   3375
            Left            =   120
            TabIndex        =   169
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
            Begin VB.CommandButton cmdCancelIdTips 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   173
               Top             =   2400
               Width           =   2655
            End
            Begin VB.CommandButton cmdUpdateTipsId 
               Caption         =   "Update both ID and Tips"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   172
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CommandButton cmdUpdateID 
               Caption         =   "Update ID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   170
               Top             =   960
               Width           =   2655
            End
            Begin VB.CommandButton cmdUpdateTips 
               Caption         =   "Update Tips"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   171
               Top             =   240
               Width           =   2655
            End
         End
      End
      Begin VB.Frame frmSalesView 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Sales Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   4680
         TabIndex        =   100
         Top             =   6120
         Width           =   3135
         Begin VB.CommandButton cmdSalesView 
            Caption         =   "View Today's Sales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.CommandButton cmdClearTran 
         Caption         =   "Clear List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   99
         Top             =   7440
         Width           =   3135
      End
   End
   Begin VB.Frame frmFinal 
      BackColor       =   &H00C0C0FF&
      Height          =   7575
      Left            =   3960
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.Frame frmBottom 
         BackColor       =   &H00C0C0FF&
         Height          =   975
         Left            =   240
         TabIndex        =   153
         Top             =   6360
         Width           =   9975
         Begin VB.CommandButton cmdCancelTran 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8040
            TabIndex        =   154
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdPrintReceipt 
         Caption         =   "Print Receipt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         TabIndex        =   66
         Top             =   360
         Width           =   2055
      End
      Begin VB.Frame frmEmpId 
         BackColor       =   &H008080FF&
         Height          =   3255
         Left            =   4200
         TabIndex        =   60
         Top             =   3000
         Width           =   2055
         Begin VB.TextBox txtEmpId 
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
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtCCTips 
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
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton cmdFinal 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   61
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblEmpId 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Employee ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label frmCCtips 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Credit Card Tips"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.PictureBox picChange 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4200
         ScaleHeight     =   1545
         ScaleWidth      =   2025
         TabIndex        =   57
         Top             =   1200
         Width           =   2055
         Begin VB.Label lblChange 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   360
            TabIndex        =   59
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblChangeAmount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame frmEmpList 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Emplyee List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   6015
         Left            =   6720
         TabIndex        =   56
         Top             =   240
         Width           =   3495
         Begin MSFlexGridLib.MSFlexGrid msgEmpList 
            Height          =   5535
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   9763
            _Version        =   393216
            Rows            =   100
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmReceiptView 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Receipt View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   6015
         Left            =   240
         TabIndex        =   137
         Top             =   240
         Width           =   3495
         Begin MSFlexGridLib.MSFlexGrid msgReceiptView 
            Height          =   5535
            Left            =   120
            TabIndex        =   138
            Top             =   360
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   9763
            _Version        =   393216
            Rows            =   100
            FixedRows       =   0
            FixedCols       =   0
            GridLines       =   0
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstRecordNum As Integer
Dim updateRecordNum As Integer
Dim tNum As Integer
Private Sub cmd10_Click()
    Let tType = 0
    Let txtTender.Text = "10.00"
End Sub

Private Sub cmd100_Click()
    Let tType = 0
    Let txtTender.Text = "100.00"
End Sub

Private Sub cmd20_Click()
    Let tType = 0
    Let txtTender.Text = "20.00"
End Sub

Private Sub cmd30_Click()
    Let tType = 0
    Let txtTender.Text = "30.00"
End Sub

Private Sub cmd40_Click()
    Let tType = 0
    Let txtTender.Text = "40.00"
End Sub

Private Sub cmd50_Click()
    Let tType = 0
    Let txtTender.Text = "50.00"
End Sub






Private Sub cmdCancel_Click()
    Call ResetSales
    Call ClearMsg(msgReceipt)
    Let lblBalance.Caption = ""
    Let txtTender.Text = ""
    Let txtTender.BackColor = &HFFFFFF
End Sub

Private Sub cmdCancelIdTips_Click()
    frmPrintViewTran.Visible = True
    frmUpdateIdTips.Visible = False
End Sub

Private Sub cmdCancelTran_Click()

    Let yn = MsgBox("Are you sure you want to Cancel this Transaction?", vbYesNo)
    If yn = 7 Then
        Exit Sub
    End If
    msgReceiptView.Clear
    msgEmpList.Clear
    Let txtCCTips.Text = ""
    Let txtEmpId.Text = ""
    frmFinal.Visible = False
    frmReceipt.Visible = True
    frmManicure.Visible = True
    frmPedicure.Visible = True
    frmThreading.Visible = True
    frmWaxing.Visible = True
    frmMassage.Visible = True
    frmTips.Visible = True
    frmOtherServices.Visible = True
    frmProduct.Visible = True
    frmManager.Visible = True
    frmDiscount.Visible = True
    frmGiftCertificate.Visible = True
    frmSpecial.Visible = True
    frmTransactionDay.Visible = True
    frmManual.Visible = True
    
    Call ResetSales
    Call ClearMsg(msgReceipt)
    Let txtTender.Text = ""
    Let lblBalance.Caption = ""
End Sub

Private Sub cmdCash_Click()
    Let tType = 0
End Sub

Private Sub cmdClear_Click()
    txtTender.Text = ""
End Sub

Private Sub cmdCloseEmployee_Click()
    frmEmployee.Visible = False
    frmReceipt.Visible = True
    frmManicure.Visible = True
    frmPedicure.Visible = True
    frmThreading.Visible = True
    frmWaxing.Visible = True
    frmMassage.Visible = True
    frmTips.Visible = True
    frmOtherServices.Visible = True
    frmProduct.Visible = True
    frmManager.Visible = True
    frmDiscount.Visible = True
    frmGiftCertificate.Visible = True
    frmTenderBack.Visible = True
    frmPackageDeal.Visible = True
    Close #2
    Call ResetSales
End Sub

Private Sub cmdClearCommision_Click()
    msgCommision.Clear
End Sub

Private Sub cmdClearTran_Click()
    Call ClearMsg(msgSalesView)
End Sub

Private Sub cmdClosing_Click()
    Closing.Show 1
End Sub

Private Sub cmdCreditCard_Click()
    Let tType = 1
    Let txtTender.Text = total - redemGCAmount
End Sub

Private Sub cmdCloseReprint_Click()
    frmPrintViewTran.Visible = True
    frmUpdateIdTips.Visible = False
    txtTips.Text = ""
    txtNewId.Text = ""
    lblTips.Visible = False
    lblNewId.Visible = False
    txtTips.Visible = False
    txtNewId.Visible = False
    frmAdjustment.Visible = False
    
    msgSalesView.Clear
    msgCommision.Clear
    comEmpList.Clear
    Let txtTransactionNum.Text = ""
    
    
    
    Call ResetSales
    
    frmReprint.Visible = False
    frmReceipt.Visible = True
    frmManicure.Visible = True
    frmPedicure.Visible = True
    frmThreading.Visible = True
    frmWaxing.Visible = True
    frmMassage.Visible = True
    frmTips.Visible = True
    frmOtherServices.Visible = True
    frmProduct.Visible = True
    frmSpecial.Visible = True
    frmManager.Visible = True
    frmDiscount.Visible = True
    frmGiftCertificate.Visible = True
    frmTransactionDay.Visible = True
    frmManual.Visible = True
End Sub

Private Sub cmdDisCancel_Click()
    lblDiscount.Caption = ""
    comDiscount.Text = "Select Right Discount"
End Sub

Private Sub cmdDone_Click()
    Dim item As salesItem
    Dim itemNum As Integer
    Dim i As Integer
        
    If balance > 0 Then
        Let txtTender.BackColor = &HFF&
        txtTender.SetFocus
        Exit Sub
    ElseIf itemIndex = 0 And giftCertNum = 0 Then
        Exit Sub
    End If
    
    Let dstr = Date
    Let tstr = Time
    Let txtTender.BackColor = &HFFFFFF
        
    frmFinal.Visible = True
    frmReceipt.Visible = False
    frmManicure.Visible = False
    frmPedicure.Visible = False
    frmThreading.Visible = False
    frmWaxing.Visible = False
    frmMassage.Visible = False
    frmTips.Visible = False
    frmOtherServices.Visible = False
    frmProduct.Visible = False
    frmManager.Visible = False
    frmDiscount.Visible = False
    frmGiftCertificate.Visible = False
    frmSpecial.Visible = False
    frmTransactionDay.Visible = False
    frmManual.Visible = False
    
    Let msgReceiptView.ColWidth(0) = 2100
    Let msgReceiptView.ColWidth(1) = 800
    Let msgReceiptView.ColAlignment(0) = 1
    Let msgReceiptView.row = 0
    Let msgReceiptView.Col = 0
    Let msgReceiptView.CellFontUnderline = True
    Let msgReceiptView.CellFontBold = True
    Let msgReceiptView.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgReceiptView.Col = 1
    Let msgReceiptView.CellFontUnderline = True
    Let msgReceiptView.CellFontBold = True
    Let msgReceiptView.Text = Format("Price", "@@@@@@@@@")
    
    For i = 1 To itemIndex
        Call ReceiptViewFormat(itemName(i), itemPrice(i), i)
    Next i
    Let lastRow = itemIndex
    For i = 1 To giftCertNum
        Let lastRow = lastRow + 1
        Call ReceiptViewFormat("Gift Certificate", giftCertArray(i), lastRow)
    Next i
    Let lastRow = lastRow + 1
    Call ReceiptViewFormat("SubTotal", subTotal, lastRow)
    Let lastRow = lastRow + 1
    Call ReceiptViewFormat("Tax", tax, lastRow)
    Let lastRow = lastRow + 1
    Call ReceiptViewFormat("Total", total, lastRow)
    
    For i = 1 To redemGCNum
        Let lastRow = lastRow + 1
        Call ReceiptViewFormat("Gift Certificate Redem", redemGCArray(i) * -1, lastRow)
    Next i
    
    If tenderAmt > 0 Then
        Let lastRow = lastRow + 1
        If tType = 1 Then
            Call ReceiptViewFormat("Credit Card", tenderAmt, lastRow)
        Else
            Call ReceiptViewFormat("Cash", tenderAmt, lastRow)
        End If
    End If
    
    Let lastRow = lastRow + 1
    Call ReceiptViewFormat("Balance", balance, lastRow)
    
    Let lblChangeAmount.Caption = Format(balance, "0.00")
    txtEmpId.SetFocus
    Let txtCCTips.Text = 0
    Dim emp As employee
    
    Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #2 Len = Len(emp)
    Let empNum = LOF(2) / Len(emp)
    
    Let msgEmpList.ColWidth(0) = 500
    Let msgEmpList.ColWidth(1) = 2500
    Let msgEmpList.ColAlignment(0) = 3
    Let msgEmpList.row = 0
    Let msgEmpList.Col = 0
    Let msgEmpList.CellFontUnderline = True
    Let msgEmpList.CellFontBold = True
    Let msgEmpList.Text = Format("ID", "@@@@@@")
    Let msgEmpList.Col = 1
    Let msgEmpList.CellFontUnderline = True
    Let msgEmpList.CellFontBold = True
    Let msgEmpList.Text = Format("Employee Name", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    
    For i = 1 To empNum
        Get #2, i, emp
        Let msgEmpList.row = i
        Let msgEmpList.Col = 0
        Let msgEmpList.Text = emp.id
        Let msgEmpList.row = i
        Let msgEmpList.Col = 1
        Let msgEmpList.Text = emp.name
    Next i
    Close #2
End Sub

Private Sub cmdExtraClear_Click()
    Let txtExtra.Text = ""
End Sub

Private Sub cmdExtraOk_Click()
    If IsNumeric(Trim(txtExtra.Text)) = False Then
        Let txtExtra.Text = "Enter Number"
        Exit Sub
    End If
    
    Let itemIndex = itemIndex + 1
    Let itemType(itemIndex) = "EXT1"
    Let itemName(itemIndex) = "Extra"
    Let itemPrice(itemIndex) = Val(Trim(txtExtra.Text))
    Let totalSales = totalSales + Val(Trim(txtExtra.Text))
    Let itemCommision(itemIndex) = 0
    Let txtExtra.Text = ""
    Call TotalDue
End Sub

Private Sub cmdFinal_Click()
    Dim item As salesItem
    Let transactionNum = transactionNum + 1
     
    Let item.name = "Emp ID"
    If IsNumeric(txtEmpId.Text) Then
        Let item.price = Val(txtEmpId.Text)
    Else
        Let item.price = 0
    End If
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 100
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
    
        
    For itemNum = 1 To itemIndex
        Let item.name = itemName(itemNum)
        Let item.price = itemPrice(itemNum)
        Let item.tranNum = transactionNum
        Let item.itemType = itemType(itemNum)
        Let item.commision = itemCommision(itemNum)
        Let item.dateStr = dstr
        Let item.timeStr = tstr
        Let recordNum = recordNum + 1
        Put #1, recordNum, item
    Next itemNum
    
    For i = 1 To giftCertNum
        Let item.name = "Gift Certificate"
        Let item.price = giftCertArray(i)
        Let item.tranNum = transactionNum
        Let item.commision = 0
        Let item.itemType = 101
        Let item.dateStr = dstr
        Let item.timeStr = tstr
        Let recordNum = recordNum + 1
        Put #1, recordNum, item
    Next i
    
    Let item.name = "SubTotal:"
    Let item.price = subTotal
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 102
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
    
    Let item.name = "Tax:"
    Let item.price = tax
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 103
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
        
            
    Let item.name = "Total:"
    Let item.price = total
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 104
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
    
    For i = 1 To redemGCNum
        Let item.name = "Gift Certificate Redem:"
        Let item.price = redemGCArray(i)
        Let item.tranNum = transactionNum
        Let item.commision = 0
        Let item.itemType = 105
        Let item.dateStr = dstr
        Let item.timeStr = tstr
        Let recordNum = recordNum + 1
        Put #1, recordNum, item
            
    Next i
    
    
    If tenderAmt > 0 Then
        If (tType = 0) Then
            Let item.name = "Cash Tendered:"
            Let item.price = tenderAmt
            Let item.tranNum = transactionNum
            Let item.commision = 0
            Let item.itemType = 106
            Let item.dateStr = dstr
            Let item.timeStr = tstr
            Let recordNum = recordNum + 1
            Put #1, recordNum, item
            
        ElseIf tType = 1 Then
            Let item.name = "Credit Card:"
            Let item.price = tenderAmt
            Let item.tranNum = transactionNum
            Let item.commision = 0
            Let item.itemType = 107
            Let item.dateStr = dstr
            Let item.timeStr = tstr
            Let recordNum = recordNum + 1
            Put #1, recordNum, item
        End If
    End If
        
    Let item.name = "Change:"
    Let item.price = balance
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 108
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
    
    Let item.name = "CC Tips"
    If IsNumeric(txtCCTips.Text) Then
        Let item.price = Val(txtCCTips.Text)
    Else
        Let item.price = 0
    End If
    Let item.tranNum = transactionNum
    Let item.commision = 0
    Let item.itemType = 109
    Let item.dateStr = dstr
    Let item.timeStr = tstr
    Let recordNum = recordNum + 1
    Put #1, recordNum, item
    
    Dim gc As gcItem
    Open defaultDir & "\nailsPOS\gc\gc.txt" For Random As #2 Len = Len(gc)
    Let gcNum = LOF(2) / Len(gc)
    For i = 1 To giftCertNum
        Let gcNum = gcNum + 1
        Let gc.amount = giftCertArray(i)
        Let gc.id = giftCertId(i)
        Let gc.status = 1
        Let gc.dstr = dstr
        Let gc.tstr = tstr
        Put #2, gcNum, gc
    Next i
    Close #2
    
    msgReceiptView.Clear
    msgEmpList.Clear
    Let txtCCTips.Text = ""
    Let txtEmpId.Text = ""
    frmFinal.Visible = False
    frmReceipt.Visible = True
    frmManicure.Visible = True
    frmPedicure.Visible = True
    frmThreading.Visible = True
    frmWaxing.Visible = True
    frmMassage.Visible = True
    frmTips.Visible = True
    frmOtherServices.Visible = True
    frmProduct.Visible = True
    frmManager.Visible = True
    frmDiscount.Visible = True
    frmGiftCertificate.Visible = True
    frmSpecial.Visible = True
    frmTransactionDay.Visible = True
    frmManual.Visible = True
    
    Call ResetSales
    Call ClearMsg(msgReceipt)
    Let txtTender.Text = ""
    Let lblBalance.Caption = ""
End Sub



Private Sub cmdMainWindow_Click()
    Call ResetSales
    Call ClearMsg(msgReceipt)
    Let lblBalance.Caption = ""
    Let txtTender.Text = ""
    POS.Show
    POS.mnuProgram.Visible = True
    Register.Hide
End Sub

Private Sub cmdManicure_Click(Index As Integer)
    
    If manicureName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 1
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 2
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 3
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 4
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 5
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 6
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 7
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 8
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 8 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 9
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 9 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 10
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 10 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 11
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 11 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAN" & 12
        Let itemName(itemIndex) = manicureName(Index + 1)
        Let itemPrice(itemIndex) = manicurePrice(Index + 1)
        Let totalSales = totalSales + manicurePrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    End If
    
    Call TotalDue
End Sub



Private Sub cmdMassage_Click(Index As Integer)
    If massageName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 1
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 2
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 3
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 4
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 5
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "MAS" & 6
        Let itemName(itemIndex) = massageName(Index + 1)
        Let itemPrice(itemIndex) = massagePrice(Index + 1)
        Let totalSales = totalSales + massagePrice(Index + 1)
        Let itemCommision(itemIndex) = massageCommision(Index + 1)
    End If
    
    Call TotalDue
End Sub

Private Sub cmdOthers_Click(Index As Integer)
    If othersName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 1
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 2
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 3
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 4
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 5
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 6
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 7
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "OTH" & 8
        Let itemName(itemIndex) = othersName(Index + 1)
        Let itemPrice(itemIndex) = othersPrice(Index + 1)
        Let totalSales = totalSales + othersPrice(Index + 1)
        Let itemCommision(itemIndex) = othersCommision(Index + 1)
    End If
End Sub

Private Sub cmdPedicure_Click(Index As Integer)
        
    If pedicureName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 1
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 2
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 3
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 4
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 5
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 6
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 7
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 8
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 8 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 9
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    ElseIf Index = 9 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "PED" & 10
        Let itemName(itemIndex) = pedicureName(Index + 1)
        Let itemPrice(itemIndex) = pedicurePrice(Index + 1)
        Let totalSales = totalSales + pedicurePrice(Index + 1)
        Let itemCommision(itemIndex) = pedicureCommision(Index + 1)
    End If
    
    Call TotalDue
End Sub

Private Sub cmdPrintCommision_Click()
    Dim s As String
    Let msgCommision.row = 0
    Let msgCommision.Col = 1
    Printer.Print Trim(msgCommision.Text)
    
    Let msgCommision.row = 1
    Let msgCommision.Col = 0
    Let s = Trim(msgCommision.Text)
    Let msgCommision.Col = 1
    Let s = s & "  " & Trim(msgCommision.Text)
    Let msgCommision.Col = 2
    Let s = s & "            " & Trim(msgCommision.Text)
    Printer.Print s
    
    For i = 2 To 30
        Let msgCommision.row = i
        Let msgCommision.Col = 0
        Let s = Trim(msgCommision.Text)
        Let msgCommision.Col = 1
        Let s = s & "  " & Trim(msgCommision.Text)
        If Trim(s) = "" Then
            Exit For
        End If
        Let msgCommision.Col = 2
        Call PrintReceipt(s, Val(Trim(msgCommision.Text)))
    Next i
    Printer.Print
    Printer.Print
    Printer.Print storeName
    Printer.EndDoc
End Sub

Private Sub cmdPrintReceipt_Click()
    Dim record As Integer
    Dim i As Integer
    
    Call PrintHead(dstr, tstr, transactionNum)
    
    For record = 1 To itemIndex
        Call PrintReceipt(RTrim(itemName(record)), itemPrice(record))
    Next record
    
    For i = 1 To giftCertNum
        Call PrintReceipt("Gift Certificate", giftCertArray(i))
    Next i
            
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("SubTotal:", subTotal)
    Call PrintReceipt("Tax:", tax)
        
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Total:", total)
    
    For i = 1 To redemGCNum
        Call PrintReceipt("Gift Certificate Redem:", redemGCArray(i) * -1)
    Next i
        
    If tenderAmt > 0 Then
        If tType = 0 Then
            Call PrintReceipt("Cash Tendered:", tenderAmt * -1)
        ElseIf tType = 1 Then
            Call PrintReceipt("Credit Card:", tenderAmt * -1)
        End If
    End If
    
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Change:", balance)
    
    Printer.Print
    Printer.Print "                   *** Thank You ***"
    Printer.EndDoc
    
    
End Sub

Private Sub cmdPrintLastTran_Click()
    Dim item As salesItem
    Dim record As Integer
    
    If recordNum < 1 Then
        Exit Sub
    End If
    
    Call ResetSales
    
    For record = recordNum To 1 Step -1
        Get #1, record, item
        If IsNumeric(Trim(item.itemType)) Then
            If Trim(item.itemType) = 100 Then
                Exit For
            ElseIf Trim(item.itemType) = 101 Then
                Let giftCertNum = giftCertNum + 1
                Let giftCertArray(giftCertNum) = item.price
            ElseIf Trim(item.itemType) = 102 Then
                Let subTotal = item.price
            ElseIf Trim(item.itemType) = 103 Then
                Let tax = item.price
            ElseIf Trim(item.itemType) = 104 Then
                Let total = item.price
            ElseIf Trim(item.itemType) = 105 Then
                Let redemGCNum = redemGCNum + 1
                Let redemGCArray(redemGCNum) = item.price
            ElseIf Trim(item.itemType) = 106 Then
                Let tType = 0
                Let tenderAmt = item.price
            ElseIf Trim(item.itemType) = 107 Then
                Let tType = 1
                Let tenderAmt = item.price
            ElseIf Trim(item.itemType) = 108 Then
                Let balance = item.price
                Let tstr = item.timeStr
                Let dstr = item.dateStr
            End If
        Else
            Let itemIndex = itemIndex + 1
            Let itemName(itemIndex) = item.name
            Let itemPrice(itemIndex) = item.price
        End If
    Next record
    
    Call PrintHead(dstr, tstr, transactionNum)
    
    If itemIndex > 0 Then
        For record = itemIndex To 1 Step -1
            Call PrintReceipt(Trim(itemName(record)), itemPrice(record))
        Next record
    End If
    
    If (giftCertNum > 0) Then
        For i = giftCertNum To 1 Step -1
            Call PrintReceipt("Gift Certificate", giftCertArray(i))
        Next i
    End If
    
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("SubTotal:", subTotal)
    Call PrintReceipt("Tax:", tax)
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Total:", total)
    
    If redemGCNum > 0 Then
        For i = redemGCNum To 1 Step -1
            Call PrintReceipt("Gift Certificate Redem:", redemGCArray(i) * -1)
        Next i
    End If
    
    If tenderAmt > 0 Then
        If tType = 0 Then
            Call PrintReceipt("Cash Tendered:", tenderAmt * -1)
        ElseIf tType = 1 Then
            Call PrintReceipt("Credit Card:", tenderAmt * -1)
        End If
    End If
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Change:", balance)
    
    Printer.Print
    Printer.Print "                   *** Thank You ***"
    Printer.EndDoc
    
End Sub

Private Sub cmdPrintTran_Click()
    Dim item As salesItem
    Dim record As Integer
    
    If IsNumeric(Trim(txtTransactionNum.Text)) Then
        Let tNum = Val(Trim(txtTransactionNum.Text))
    Else
        txtTransactionNum.Text = "Enter a Number."
        Exit Sub
    End If
    
    If tNum > transactionNum Or tNum < 1 Then
        txtTransactionNum.Text = "Not Found."
        Exit Sub
    End If
    
    
    
    Call ResetSales
       
    For record = 1 To recordNum
        Get #1, record, item
        If tNum = item.tranNum Then
            If IsNumeric(Trim(item.itemType)) Then
                If Trim(item.itemType) = 101 Then
                    Let giftCertNum = giftCertNum + 1
                    Let giftCertArray(giftCertNum) = item.price
                ElseIf Trim(item.itemType) = 102 Then
                    Let subTotal = item.price
                ElseIf Trim(item.itemType) = 103 Then
                    Let tax = item.price
                ElseIf Trim(item.itemType) = 104 Then
                    Let total = item.price
                ElseIf Trim(item.itemType) = 105 Then
                    Let redemGCNum = redemGCNum + 1
                    Let redemGCArray(redemGCNum) = item.price
                ElseIf Trim(item.itemType) = 106 Then
                    Let tType = 0
                    Let tenderAmt = item.price
                ElseIf Trim(item.itemType) = 107 Then
                    Let tType = 1
                    Let tenderAmt = item.price
                ElseIf Trim(item.itemType) = 108 Then
                    Let balance = item.price
                    Let tstr = item.timeStr
                    Let dstr = item.dateStr
                ElseIf Trim(item.itemType) = 109 Then
                    Exit For
                End If
            Else
                Let itemIndex = itemIndex + 1
                Let itemName(itemIndex) = item.name
                Let itemPrice(itemIndex) = item.price
            End If
        End If
    Next record
           
    Call PrintHead(dstr, tstr, transactionNum)
    
    For record = 1 To itemIndex
        Call PrintReceipt(Trim(itemName(record)), itemPrice(record))
    Next record
    
    For i = 1 To giftCertNum
        Call PrintReceipt("Gift Certificate", giftCertArray(i))
    Next i
            
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("SubTotal:", subTotal)
    Call PrintReceipt("Tax:", tax)
        
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Total:", total)
    
    For i = 1 To redemGCNum
        Call PrintReceipt("Gift Certificate Redem:", redemGCArray(i) * -1)
    Next i
        
    If tenderAmt > 0 Then
        If tType = 0 Then
            Call PrintReceipt("Cash Tendered:", tenderAmt * -1)
        ElseIf tType = 1 Then
            Call PrintReceipt("Credit Card:", tenderAmt * -1)
        End If
    End If
    
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Change:", balance)
    
    Printer.Print
    Printer.Print "                   *** Thank You ***"
    Printer.EndDoc
    txtTransactionNum.Text = ""
    
End Sub


Private Sub cmdProductOK_Click()
    
    Dim pType As Integer
    
    Let pType = comProductList.ListIndex
    
    If pType < 0 Then
        Let comProductList.Text = "First Select Product Here."
        Exit Sub
    End If
    
    Let itemIndex = itemIndex + 1
    Let itemType(itemIndex) = "PRO" & pType + 1
    Let itemName(itemIndex) = productName(pType + 1)
    Let itemPrice(itemIndex) = productPrice(pType + 1)
    Let totalSales = totalSales + productPrice(pType + 1)
    Let itemCommision(itemIndex) = 0
    
    Call TotalDue
    
    comProductList.Text = "Select Right Product"
End Sub

Private Sub cmdRedemGC_Click()
    Rem Let GiftCertificate.frmRedemGC.Visible = True
    Rem GiftCertificate.Show 1
    Dim gcRedem As Single
    If itemIndex = 0 Then
        Exit Sub
    End If
    Let gcRedem = Val(InputBox("Enter the Amount of the Gift Certificate", "Gift Certificate"))
    If gcRedem > 0 Then
        Let gcRedem = Format(gcRedem, "0.00")
        Let redemGCNum = redemGCNum + 1
        Let redemGCArray(redemGCNum) = gcRedem
        Let redemGCAmount = redemGCAmount + gcRedem
        Call TotalDue
    End If
End Sub

Private Sub cmdReview_Click()
    
    Call ResetSales
    Let lblBalance.Caption = ""
    Let txtTender.Text = ""
    Call ClearMsg(msgReceipt)
    Let txtTender.BackColor = &HFFFFFF
    
    frmReprint.Visible = True
    frmReceipt.Visible = False
    frmManicure.Visible = False
    frmPedicure.Visible = False
    frmThreading.Visible = False
    frmWaxing.Visible = False
    frmMassage.Visible = False
    frmTips.Visible = False
    frmOtherServices.Visible = False
    frmProduct.Visible = False
    frmSpecial.Visible = False
    frmManager.Visible = False
    frmDiscount.Visible = False
    frmGiftCertificate.Visible = False
    frmTransactionDay.Visible = False
    frmManual.Visible = False
    
      
    Dim emp As employee
    
    Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #2 Len = Len(emp)
    Let empNum = LOF(2) / Len(emp)
    
    For i = 1 To empNum
        Get #2, i, emp
        comEmpList.AddItem emp.id & "  " & emp.name
    Next i
    Close #2
    
    Let msgCommision.ColWidth(0) = 500
    Let msgCommision.ColWidth(1) = 2500
    Let msgCommision.ColWidth(2) = 800
    Let msgCommision.ColAlignment(1) = 1
    Let msgCommision.ColAlignment(0) = 3
    
    Let msgSalesView.ColWidth(0) = 600
    Let msgSalesView.ColWidth(1) = 2400
    Let msgSalesView.ColWidth(2) = 800
    Let msgSalesView.ColAlignment(1) = 1
    Let msgSalesView.ColAlignment(0) = 3
    Let msgSalesView.row = 0
    Let msgSalesView.Col = 0
    Let msgSalesView.CellFontUnderline = True
    Let msgSalesView.CellFontBold = True
    Let msgSalesView.Text = Format("T. No", "@@@@@")
    Let msgSalesView.Col = 1
    Let msgSalesView.CellFontUnderline = True
    Let msgSalesView.CellFontBold = True
    Let msgSalesView.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgSalesView.Col = 2
    Let msgSalesView.CellFontUnderline = True
    Let msgSalesView.CellFontBold = True
    Let msgSalesView.Text = Format("Price", "@@@@@@@@@")
End Sub

Private Sub cmdSalesView_Click()
    Dim Password As String
    Dim subTotal As Single
    Dim tSales As Single
    Dim tTax As Single
    Dim tGC As Single
    Dim tCash As Single
    Dim tCC As Single
    Dim tGCR As Single
    Dim tCommision As Single
    Dim tTips  As Single
    
    Dim item As salesItem
    Dim record As Integer
        
        Let subTotal = 0
        Let tSales = 0
        Let tTax = 0
        Let tGC = 0
        Let tGCR = 0
        Let tCash = 0
        Let tCC = 0
        Let tCommision = 0
        Let tTips = 0
        
        For record = 1 To recordNum
            Get #1, record, item
            Let tCommision = tCommision + item.commision
            
            If Trim(item.itemType) = 101 Then
                Let tGC = tGC + item.price
            ElseIf Trim(item.itemType) = 102 Then
                Let subTotal = subTotal + item.price
            ElseIf Trim(item.itemType) = 103 Then
                Let tTax = tTax + item.price
            ElseIf Trim(item.itemType) = 104 Then
                Let tSales = tSales + item.price
            ElseIf Trim(item.itemType) = 105 Then
                Let tGCR = tGCR + item.price
            ElseIf Trim(item.itemType) = 106 Then
              
            ElseIf Trim(item.itemType) = 107 Then
                Let tCC = tCC + item.price
            ElseIf Trim(item.itemType) = 109 Then
                Let tTips = tTips + item.price
            End If
        Next record
        msgCommision.Clear
        msgCommision.row = 0
        msgCommision.Col = 1
        msgCommision.CellFontBold = True
        msgCommision.Text = "Today's Sales Summary:"
        msgCommision.row = 1
        msgCommision.Col = 1
        msgCommision.Text = "----------------------------------------------------"
        msgCommision.Col = 2
        msgCommision.Text = "--------------------------------------"
        msgCommision.row = 2
        msgCommision.Col = 1
        msgCommision.Text = "Total Gross Sales: "
        msgCommision.Col = 2
        msgCommision.Text = Format(subTotal, "0.00")
        msgCommision.row = 3
        msgCommision.Col = 1
        msgCommision.Text = "Total GiftCertificate sold: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tGC, "0.00")
        msgCommision.row = 4
        msgCommision.Col = 1
        msgCommision.Text = "Total Sales Tax: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tTax, "0.00")
        msgCommision.row = 5
        msgCommision.Col = 1
        msgCommision.Text = "Total GiftCertificate Redemed: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tGCR, "0.00")
        msgCommision.row = 6
        msgCommision.Col = 1
        msgCommision.Text = "Total Commision: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tCommision, "0.00")
        msgCommision.row = 7
        msgCommision.Col = 1
        msgCommision.CellFontBold = True
        msgCommision.Text = "Total Amount in Register: "
        msgCommision.Col = 2
        Let net = tSales - tGCR - tCommision
        msgCommision.Text = Format(net, "0.00")
        msgCommision.row = 8
        msgCommision.Col = 1
        msgCommision.CellFontBold = True
        msgCommision.Text = "Total Cash in register: "
        msgCommision.Col = 2
        msgCommision.Text = Format(net - tCC - tTips, "0.00")
        msgCommision.row = 9
        msgCommision.Col = 1
        msgCommision.CellFontBold = True
        msgCommision.Text = "Total Credit Card: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tCC + tTips, "0.00")
        msgCommision.row = 10
        msgCommision.Col = 1
        msgCommision.Text = "Total Credit Card tips: "
        msgCommision.Col = 2
        msgCommision.Text = Format(tTips, "0.00")
    
End Sub

Private Sub cmdSellGC_Click()
           
    Let GiftCertificate.msgGiftcertificate.ColWidth(0) = 2000
    Let GiftCertificate.msgGiftcertificate.ColWidth(1) = 1000
    Let GiftCertificate.msgGiftcertificate.ColAlignment(0) = 1
    Let GiftCertificate.frmSellGiftcertificate.Visible = True
    GiftCertificate.Show 1
    For i = 1 To 20
        If gcAmount(i) > 0 Then
            Let giftCertNum = giftCertNum + 1
            Let giftCertArray(giftCertNum) = gcAmount(i)
            Let giftCertId(giftCertNum) = gcId(i)
            Let giftCertSum = giftCertSum + gcAmount(i)
            Let gcAmount(i) = 0
            Let gcId(i) = 0
        Else
            Exit For
        End If
    Next i
    Call TotalDue
    
End Sub


Private Sub cmdShowCommision_Click()
    Dim item As salesItem
    Dim r As Integer
    Dim commision As Single
    Dim id As Integer
    Dim row As Integer
    Dim match As Boolean
        
    Let match = False
    Let row = 1
    Let commision = 0
    
    msgCommision.Clear
    Let id = comEmpList.ListIndex + 1
    
    Let msgCommision.row = 0
    Let msgCommision.Col = 1
    Let msgCommision.CellFontBold = True
    If id = 0 Then
        Let msgCommision.Text = "Select Employee First"
        Exit Sub
    End If
    
    Let msgCommision.Text = comEmpList.Text
    
    
    Let msgCommision.row = 1
    Let msgCommision.Col = 0
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("TN", "@@@")
    Let msgCommision.Col = 1
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgCommision.Col = 2
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("Amount", "@@@@@@@")
    
    For r = 1 To recordNum
        Get #1, r, item
        If Trim(item.itemType) = 100 Then
            If id = item.price Then
                Let match = True
            Else
                Let r = r + 7
            End If
        ElseIf match = True Then
            If Trim(item.itemType) = 109 Then
                If item.price > 0 Then
                    Let row = row + 1
                    Let msgCommision.row = row
                    Let msgCommision.Col = 0
                    Let msgCommision.Text = item.tranNum
                    Let msgCommision.Col = 1
                    Let msgCommision.Text = "Credit Card Tips"
                    Let msgCommision.Col = 2
                    Let msgCommision.Text = Format(item.price, "0.00")
                    Let commision = commision + item.price
                End If
                Let match = False
            ElseIf item.commision > 0 Then
                Let row = row + 1
                Let msgCommision.row = row
                Let msgCommision.Col = 0
                Let msgCommision.Text = item.tranNum
                Let msgCommision.Col = 1
                Let msgCommision.Text = Trim(item.name) & " Commision"
                Let msgCommision.Col = 2
                Let msgCommision.Text = Format(item.commision, "0.00")
                Let commision = commision + item.commision
            End If
        End If
    Next r
    If commision = 0 Then
        Let msgCommision.row = row + 1
        Let msgCommision.Col = 1
        Let msgCommision.Text = "No tips and commision."
        Exit Sub
    End If
    Let msgCommision.row = row + 1
    Let msgCommision.Col = 1
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = "Total"
    Let msgCommision.Col = 2
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format(commision, "0.00")
End Sub

Private Sub cmdSpecial_Click(Index As Integer)
    If specialName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 1
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 2
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 3
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 4
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 5
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "SPE" & 6
        Let itemName(itemIndex) = specialName(Index + 1)
        Let itemPrice(itemIndex) = specialPrice(Index + 1)
        Let totalSales = totalSales + specialPrice(Index + 1)
        Let itemCommision(itemIndex) = specialCommision(Index + 1)
    End If
    Call TotalDue
End Sub

Private Sub cmdThreading_Click(Index As Integer)
        
    If threadingName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 1
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 2
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 3
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 4
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 5
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 6
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 7
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "THR" & 8
        Let itemName(itemIndex) = threadingName(Index + 1)
        Let itemPrice(itemIndex) = threadingPrice(Index + 1)
        Let totalSales = totalSales + threadingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    End If
    
    Call TotalDue
End Sub

Private Sub cmdTicket_Click()
    ServiceTicket.Show 1
End Sub

Private Sub cmdTipAdjustment_Click()
    Dim item As salesItem
    Dim record As Integer
    
    Dim row As Integer
    
    If IsNumeric(Trim(txtTransactionNum.Text)) Then
        Let tNum = Val(Trim(txtTransactionNum.Text))
    Else
        Let txtTransactionNum.Text = "Enter a Number."
        Exit Sub
    End If
    
    If tNum > transactionNum Or tNum < 1 Then
        txtTransactionNum.Text = "Not Found."
        Exit Sub
    End If
    
    Call ClearMsg(msgSalesView)
    
    Let row = 0
    For record = 1 To recordNum
        Get #1, record, item
        If tNum = item.tranNum Then
            Let row = row + 1
            msgSalesView.row = row
            
            If row = 1 Then
                Let dstr = item.dateStr
                Let tstr = item.timeStr
                Let firstRecordNum = record
                
                msgSalesView.Col = 0
                msgSalesView.Text = item.tranNum
                msgSalesView.Col = 1
                msgSalesView.Text = item.name
                msgSalesView.Col = 2
                msgSalesView.Text = Format(item.price, "0.00")
            Else
                msgSalesView.Col = 1
                msgSalesView.Text = item.name
                msgSalesView.Col = 2
                msgSalesView.Text = Format(item.price, "0.00")
            End If
            If Trim(item.itemType) = 109 Then
                Let updateRecordNum = record
                Exit For
            End If
        End If
    Next record
    Let txtTransactionNum.Text = ""
    frmPrintViewTran.Visible = False
    frmUpdateIdTips.Visible = True
    
End Sub

Private Sub cmdTips_Click(Index As Integer)
    If tipsName(Index + 1) = "" Then
        Exit Sub
    End If
        
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 1
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 2
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 3
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 4
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 5
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 6
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 7
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 8
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 8 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 9
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 9 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "TIP" & 10
        Let itemName(itemIndex) = tipsName(Index + 1)
        Let itemPrice(itemIndex) = tipsPrice(Index + 1)
        Let totalSales = totalSales + tipsPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    End If
    
    Call TotalDue
End Sub

Private Sub cmdTotal_Click()
    Call TotalDue
End Sub

Private Sub cmdMgrWindowClose_Click()
    frmReceipt.Visible = True
    frmMgrWindow.Visible = False
    listTransaction.Clear
    Call ResetSales
End Sub


Private Sub cmdTipsCancel_Click()
    Let txtTips.Text = ""
    Let txtNewId.Text = ""
    lblTips.Visible = False
    lblNewId.Visible = False
    txtTips.Visible = False
    txtNewId.Visible = False
    Call ClearMsg(msgSalesView)
    frmPrintViewTran.Visible = True
    frmAdjustment.Visible = False
    
End Sub

Private Sub cmdTipsOK_Click()
    Dim item As salesItem
    
    If txtTips.Visible = True Then
        Let item.name = "CC Tips"
        Let item.price = Val(Trim(txtTips.Text))
        Let item.tranNum = tNum
        Let item.commision = 0
        Let item.itemType = 109
        Let item.dateStr = dstr
        Let item.timeStr = tstr
        Put #1, updateRecordNum, item
    End If
    
    If txtNewId.Visible = True Then
        Let item.name = "Emp ID"
        
        If IsNumeric(txtNewId.Text) = True Then
            Let item.price = Val(Trim(txtNewId.Text))
        Else
            
        End If
        Let item.tranNum = tNum
        Let item.commision = 0
        Let item.itemType = 100
        Let item.dateStr = dstr
        Let item.timeStr = tstr
        Put #1, firstRecordNum, item
    End If
    
    Call ClearMsg(msgSalesView)
    
    Let row = 0
    For record = firstRecordNum To updateRecordNum
        Get #1, record, item
        Let row = row + 1
        msgSalesView.row = row
            
        If row = 1 Then
            msgSalesView.Col = 0
            msgSalesView.Text = item.tranNum
            msgSalesView.Col = 1
            msgSalesView.Text = item.name
            msgSalesView.Col = 2
            msgSalesView.Text = Format(item.price, "0.00")
        Else
            msgSalesView.Col = 1
            msgSalesView.Text = item.name
            msgSalesView.Col = 2
            msgSalesView.Text = Format(item.price, "0.00")
        End If
    Next record
    Let txtTips.Text = ""
    Let txtNewId.Text = ""
    lblTips.Visible = False
    lblNewId.Visible = False
    txtTips.Visible = False
    txtNewId.Visible = False
    
    frmAdjustment.Visible = False
    frmPrintViewTran.Visible = True
    
End Sub

Private Sub cmdUpdateID_Click()
    frmUpdateIdTips.Visible = False
    frmAdjustment.Visible = True
    Let lblNewId.Visible = True
    Let txtNewId.Visible = True
End Sub

Private Sub cmdUpdateTips_Click()
    frmUpdateIdTips.Visible = False
    frmAdjustment.Visible = True
    Let lblTips.Visible = True
    Let txtTips.Visible = True
End Sub

Private Sub cmdUpdateTipsId_Click()
    frmUpdateIdTips.Visible = False
    frmAdjustment.Visible = True
    lblNewId.Visible = True
    txtNewId.Visible = True
    lblTips.Visible = True
    txtTips.Visible = True
    
End Sub

Private Sub cmdViewAllTran_Click()
    Dim item As salesItem
    Dim record As Integer
    Dim tran As Integer
    
    Call ClearMsg(msgSalesView)
    For record = 1 To recordNum
        Get #1, record, item
        If tran < item.tranNum Then
            msgSalesView.row = record
            msgSalesView.Col = 0
            msgSalesView.Text = item.tranNum
            msgSalesView.Col = 1
            msgSalesView.Text = Trim(item.name)
            msgSalesView.Col = 2
            msgSalesView.Text = Format(item.price, "0.00")
            tran = item.tranNum
        Else
            msgSalesView.row = record
            msgSalesView.Col = 1
            msgSalesView.Text = Trim(item.name)
            msgSalesView.Col = 2
            msgSalesView.Text = Format(item.price, "0.00")
        End If
    Next record
    
End Sub

Private Sub cmdViewEmpTran_Click()
    Dim item As salesItem
    Dim row As Integer
    Dim r As Integer
    Dim id As Integer
    
    Let row = 1
    msgCommision.Clear
    Let id = comEmpList.ListIndex + 1
    
    Let msgCommision.row = 0
    Let msgCommision.Col = 1
    Let msgCommision.CellFontBold = True
    If id = 0 Then
        Let msgCommision.Text = "Select Employee First"
        Exit Sub
    End If
    
    Let msgCommision.Text = comEmpList.Text
    
    
    Let msgCommision.row = 1
    Let msgCommision.Col = 0
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("TN", "@@@")
    Let msgCommision.Col = 1
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgCommision.Col = 2
    Let msgCommision.CellFontUnderline = True
    Let msgCommision.CellFontBold = True
    Let msgCommision.Text = Format("Price", "@@@@@@@")
    
    For r = 1 To recordNum
        Get #1, r, item
        If Trim(item.itemType) = 100 Then
            If id = item.price Then
                Let match = True
                Let row = row + 1
                Let msgCommision.row = row
                Let msgCommision.Col = 0
                Let msgCommision.Text = item.tranNum
            Else
                Let r = r + 7
            End If
        ElseIf match = True Then
            Let msgCommision.Col = 1
            Let msgCommision.Text = Trim(item.name)
            Let msgCommision.Col = 2
            Let msgCommision.Text = Format(item.price, "0.00")
            
            If Trim(item.itemType) = 109 Then
                Let match = False
            End If
            Let row = row + 1
            Let msgCommision.row = row
        End If
    Next r
End Sub

Private Sub cmdViewTran_Click()
    Dim item As salesItem
    Dim record As Integer
    Dim row As Integer
    
    If IsNumeric(Trim(txtTransactionNum.Text)) Then
        Let tNum = Val(Trim(txtTransactionNum.Text))
    Else
        Let txtTransactionNum.Text = "Enter a Number."
        Exit Sub
    End If
    
    If tNum > transactionNum Or tNum < 1 Then
        txtTransactionNum.Text = "Not Found."
        Exit Sub
    End If
    
    Call ViewTran(tNum)
    
    txtTransactionNum.Text = ""
End Sub

Private Sub cmdVoid_Click()
    Dim item As salesItem
    Dim voidItem As salesItem
    Dim record As Integer
    
    
    If IsNumeric(Trim(txtTransactionNum.Text)) Then
        Let tNum = Val(Trim(txtTransactionNum.Text))
    Else
        Let txtTransactionNum.Text = "Enter a Number."
        Exit Sub
    End If
    
    If tNum > transactionNum Or tNum < 1 Then
        txtTransactionNum.Text = "Not Found."
        Exit Sub
    End If
    
    Call ViewTran(tNum)
            
    If ChkPassword() = False Then
        Exit Sub
    End If
    
    Call ResetSales
    
              
    For record = 1 To recordNum
        Get #1, record, item
        If tNum = item.tranNum Then
            If IsNumeric(Trim(item.itemType)) Then
                If Trim(item.itemType) = 100 Then
                    Let voidItem.name = "Emp ID-Void"
                    Let voidItem.price = item.price
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 100
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 101 Then
                    Let voidItem.name = "GiftCertificate-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 101
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 102 Then
                    Let voidItem.name = "SubTotal-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 102
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 103 Then
                    Let voidItem.name = "Tax-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 103
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 104 Then
                    Let voidItem.name = "Total-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 104
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 105 Then
                    Let voidItem.name = "GC Redem-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 105
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 106 Then
                    Let voidItem.name = "Cash-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 106
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 107 Then
                    Let voidItem.name = "CreditCard-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 107
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 108 Then
                    Let voidItem.name = "Balance-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 108
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                ElseIf Trim(item.itemType) = 109 Then
                    Let voidItem.name = "CC Tips-Void"
                    Let voidItem.price = 0
                    Let voidItem.tranNum = tNum
                    Let voidItem.commision = 0
                    Let voidItem.itemType = 109
                    Let voidItem.dateStr = item.dateStr
                    Let voidItem.timeStr = item.timeStr
                    Put #1, record, voidItem
                    Exit For
                End If
            Else
                Let voidItem.name = Trim(item.name) & "...Void"
                Let voidItem.price = 0
                Let voidItem.tranNum = tNum
                Let voidItem.commision = 0
                Let voidItem.itemType = item.itemType
                Let voidItem.dateStr = item.dateStr
                Let voidItem.timeStr = item.timeStr
                Put #1, record, voidItem
            End If
        End If
    Next record
         
    Call ViewTran(tNum)
    txtTransactionNum.Text = ""
End Sub

Private Sub cmdWaxing_Click(Index As Integer)
    If waxingName(Index + 1) = "" Then
        Exit Sub
    End If
    
    If Index = 0 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 1
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 1 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 2
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 2 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 3
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 3 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 4
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 4 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 5
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 5 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 6
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 6 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 7
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 7 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 8
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 8 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 9
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 9 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 10
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 10 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 11
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 11 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 12
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 12 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 13
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    ElseIf Index = 13 Then
        Let itemIndex = itemIndex + 1
        Let itemType(itemIndex) = "WAX" & 14
        Let itemName(itemIndex) = waxingName(Index + 1)
        Let itemPrice(itemIndex) = waxingPrice(Index + 1)
        Let totalSales = totalSales + waxingPrice(Index + 1)
        Let itemCommision(itemIndex) = 0
    End If
    
    Call TotalDue
End Sub

Public Sub TotalDue()
    Dim i As Integer
    
    Call ClearMsg(msgReceipt)
    Let lblBalance.Caption = ""
    
    If itemIndex = 0 And giftCertNum = 0 Then
        Exit Sub
    End If
    
    For i = 1 To itemIndex
        Call PrintFormat(itemName(i), itemPrice(i), i)
    Next i
    Let lastRow = itemIndex
        
    If giftCertNum > 0 Then
        For i = 1 To giftCertNum
            Let lastRow = lastRow + 1
            Call PrintFormat("Gift Certificate", giftCertArray(i), lastRow)
        Next i
    End If
        
    Let subTotal = totalSales + giftCertSum
    Let lastRow = lastRow + 1
    Call PrintFormat("SubTotal", Format(subTotal, "0.00"), lastRow)
    Let tax = Format(totalSales * 0.045, "0.00")
    If tax > 0 Then
        Let lastRow = lastRow + 1
        Call PrintFormat("Tax", tax, lastRow)
    End If
    
    Let lastRow = lastRow + 1
    Let total = Format(subTotal + tax, "0.00")
    Call PrintFormat("Total", total, lastRow)
        
    If redemGCNum > 0 Then
        For i = 1 To redemGCNum
            Let lastRow = lastRow + 1
            Call PrintFormat("Gift Certificate Redem", redemGCArray(i) * -1, lastRow)
        Next i
    End If
    
    If tenderAmt > 0 Then
        Let lastRow = lastRow + 1
        If tType = 1 Then
            Call PrintFormat("Credit Card", tenderAmt * -1, lastRow)
        Else
            Call PrintFormat("Cash", tenderAmt * -1, lastRow)
        End If
    End If
    Let lastRow = lastRow + 1
    Let balance = Format(total - redemGCAmount - tenderAmt, "0.00")
    Let lblBalance.Caption = "Balance: " & Format(balance, "0.00")
    
    txtTender.SetFocus
   
End Sub

Private Sub Form_Load()
       
    Dim item As salesItem
    Dim dateStr As String
    
    dateStr = Format(Date, "mmddyyyy")
    Let fileName = "sale" & dateStr & ".txt"
    
    If Dir(defaultDir & "\nailsPOS\sales\closed" & fileName) = "" Then
        
        Open defaultDir & "\nailsPOS\sales\" & fileName For Random As #1 Len = Len(item)
        Let recordNum = LOF(1) / Len(item)
        If recordNum > 0 Then
            Get #1, recordNum, item
            Let transactionNum = item.tranNum
        Else
            Let recordNum = 0
            Let transactionNum = 0
        End If
        
        Let lblToday.Caption = "Today is: " & Format(Date, "dddd, mmmm dd, yyyy")
        Let msgReceipt.ColWidth(0) = 2150
        Let msgReceipt.ColWidth(1) = 750
        Let msgReceipt.ColAlignment(0) = 1
        Let msgReceipt.row = 0
        Let msgReceipt.Col = 0
        Let msgReceipt.CellFontUnderline = True
        Let msgReceipt.CellFontBold = True
        Let msgReceipt.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        Let msgReceipt.Col = 1
        Let msgReceipt.CellFontUnderline = True
        Let msgReceipt.CellFontBold = True
        Let msgReceipt.Text = Format("Price", "@@@@@@@")
    End If
End Sub

Private Sub msgEmpList_DblClick()
    Let msgEmpList.Col = 0
    Let txtEmpId.Text = Trim(msgEmpList.Text)
End Sub

Private Sub msgReceipt_DblClick()
    Dim row As Integer
    Dim s As String
    
    Let row = msgReceipt.row
    Let msgReceipt.Col = 0
    Let s = Trim(msgReceipt.Text)
    
    If row = 0 Or s = "" Or s = "SubTotal" Or s = "Tax" Or s = "Total" Then
        Exit Sub
    ElseIf s = "Gift Certificate" Then
        Let giftCertSum = giftCertSum - giftCertArray(row - itemIndex)
        For i = row - itemIndex To giftCertNum
            Let giftCertArray(i) = giftCertArray(i + 1)
            Let giftCertId(i) = giftCertId(i + 1)
        Next i
        Let giftCertNum = giftCertNum - 1
    ElseIf s = "Cash" Or s = "Credit Card" Then
        Let tenderAmt = 0
        Let tType = 0
        Let txtTender.Text = ""
    ElseIf s = "Gift Certificate Redem" Then
        Let j = itemIndex + giftCertNum + 3
        Let redemGCAmount = redemGCAmount - redemGCArray(row - j)
        For i = row - j To redemGCNum
            Let redemGCArray(i) = redemGCArray(i + 1)
        Next i
        Let redemGCNum = redemGCNum - 1
    ElseIf row <= itemIndex Then
        Let totalSales = totalSales - itemPrice(row)
        For i = row To itemIndex
            Let itemType(i) = itemType(i + 1)
            Let itemName(i) = itemName(i + 1)
            Let itemPrice(i) = itemPrice(i + 1)
            Let itemCommision(i) = itemCommision(i + 1)
        Next i
        Let itemIndex = itemIndex - 1
    End If
            
    If itemIndex + giftCertNum = 0 Then
        Call ResetSales
    End If
    
    Call TotalDue
End Sub


Private Sub tmrCurrentTime_Timer()
    Let lblTime.Caption = Format(Time, "h:nn:ss AM/PM")
End Sub

Private Sub txtCCTips_GotFocus()
    Let txtCCTips.Text = ""
End Sub


Private Sub txtExtra_GotFocus()
    Let txtExtra.Text = ""
End Sub

Private Sub txtNewId_GotFocus()
    Let txtNewId.Text = ""
End Sub

Private Sub txtTender_Change()
    Let txtTender.BackColor = &HFFFFFF
    Let tenderAmt = Val(Trim(txtTender.Text))
    Call TotalDue
End Sub

Public Sub PrintFormat(str As String, num As Single, row As Integer)
    If str = "SubTotal" Or str = "Total" Or str = "Tax" Then
        Let msgReceipt.row = row
        Let msgReceipt.Col = 0
        msgReceipt.CellFontBold = True
        Let msgReceipt.Text = str
        Let msgReceipt.Col = 1
        Let msgReceipt.Text = Format(num, "0.00")
    Else
        Let msgReceipt.row = row
        Let msgReceipt.Col = 0
        msgReceipt.CellFontBold = False
        Let msgReceipt.Text = str
        Let msgReceipt.Col = 1
        Let msgReceipt.Text = Format(num, "0.00")
    End If
End Sub


Private Sub txtTender_KeyPress(KeyAscii As Integer)
    If IsNumeric(KeyAscii) Then
        Let tType = 0
    End If
End Sub

Public Sub ReceiptViewFormat(str As String, num As Single, row As Integer)
    If str = "SubTotal" Or str = "Total" Or str = "Tax" Then
        Let msgReceiptView.row = row
        Let msgReceiptView.Col = 0
        msgReceiptView.CellFontBold = True
        Let msgReceiptView.Text = str
        Let msgReceiptView.Col = 1
        Let msgReceiptView.Text = Format(num, "0.00")
    Else
        Let msgReceiptView.row = row
        Let msgReceiptView.Col = 0
        msgReceiptView.CellFontBold = False
        Let msgReceiptView.Text = str
        Let msgReceiptView.Col = 1
        Let msgReceiptView.Text = Format(num, "0.00")
    End If
End Sub

Public Sub ViewTran(tNum As Integer)
    Dim row As Integer
    Dim record As Integer
    Dim item As salesItem
    
    Call ClearMsg(msgSalesView)
    
    Let row = 0
    For record = 1 To recordNum
        Get #1, record, item
        If tNum = item.tranNum Then
            Let row = row + 1
            msgSalesView.row = row
            
            If row = 1 Then
                msgSalesView.Col = 0
                msgSalesView.Text = item.tranNum
                msgSalesView.Col = 1
                msgSalesView.Text = item.name
                msgSalesView.Col = 2
                msgSalesView.Text = Format(item.price, "0.00")
            Else
                msgSalesView.Col = 1
                msgSalesView.Text = item.name
                msgSalesView.Col = 2
                msgSalesView.Text = Format(item.price, "0.00")
            End If
            
            If Trim(item.itemType) = 109 Then
                Exit For
            End If
        End If
    Next record
End Sub



Private Sub txtTips_GotFocus()
    Let txtTips.Text = ""
End Sub

Private Sub txtTransactionNum_GotFocus()
    Let txtTransactionNum.Text = ""
End Sub
