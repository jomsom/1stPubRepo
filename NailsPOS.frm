VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form POS 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12150
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleMode       =   0  'User
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDisplay 
      BackColor       =   &H00C0C0FF&
      Height          =   8775
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   11895
      Begin VB.PictureBox picRegister 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   9000
         ScaleHeight     =   3105
         ScaleWidth      =   2505
         TabIndex        =   99
         Top             =   600
         Width           =   2535
         Begin VB.Timer tmrCurrentTime 
            Interval        =   100
            Left            =   360
            Top             =   1200
         End
         Begin VB.CommandButton cmdRegister 
            Caption         =   "Register"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   100
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblToday 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   102
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   735
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame frmLogIn 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Manager Window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1935
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   2415
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton cmdLogin 
            Caption         =   "Log In"
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
            TabIndex        =   28
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdLogout 
            Caption         =   "Log Out"
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
            TabIndex        =   29
            Top             =   1200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lblPassword 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame frmManagerTask 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Manager Task"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   4095
         Left            =   360
         TabIndex        =   31
         Top             =   4320
         Visible         =   0   'False
         Width           =   11175
         Begin VB.CommandButton cmdSchedule 
            Caption         =   "Schedule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   2520
            TabIndex        =   57
            Top             =   2640
            Width           =   2415
         End
         Begin VB.CommandButton cmdReview 
            Caption         =   "Review"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   2520
            TabIndex        =   39
            Top             =   1560
            Width           =   2415
         End
         Begin VB.CommandButton cmdStoreProfile 
            Caption         =   "Store Profile"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   2520
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   38
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton cmdSetup 
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   240
            TabIndex        =   34
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CommandButton cmdMenuFile 
            Caption         =   "Menu File"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CommandButton cmdEmployeeFile 
            Caption         =   "Employee File"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame frmManagerTaskBk 
         BackColor       =   &H00C0C0FF&
         Height          =   4095
         Left            =   360
         TabIndex        =   36
         Top             =   4320
         Width           =   11175
      End
   End
   Begin VB.Frame frmStoreProfile 
      Caption         =   "Store Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   960
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame frmProfileView 
         Caption         =   "Current Profile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   6720
         TabIndex        =   81
         Top             =   360
         Width           =   3015
         Begin VB.PictureBox picProfileView 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   2745
            TabIndex        =   82
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdProfileClose 
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
         Height          =   615
         Left            =   3960
         TabIndex        =   56
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdProfileUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   55
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtWeb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   54
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox txtZip 
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
         Left            =   4680
         TabIndex        =   52
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtPhone 
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
         Left            =   1320
         TabIndex        =   50
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtState 
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
         Left            =   1320
         TabIndex        =   48
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtCity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   46
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtStreet 
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
         Left            =   1320
         TabIndex        =   44
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtStoreName 
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
         Left            =   1320
         TabIndex        =   42
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblWeb 
         Caption         =   "Web"
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
         Left            =   480
         TabIndex        =   53
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
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
         Left            =   4080
         TabIndex        =   51
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
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
         Left            =   480
         TabIndex        =   49
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblState 
         Caption         =   "State"
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
         Left            =   480
         TabIndex        =   47
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblCity 
         Caption         =   "City"
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
         Left            =   480
         TabIndex        =   45
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblAddress 
         Caption         =   "Street"
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
         Left            =   480
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblStoreName 
         Caption         =   "Name"
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
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame frmSetup 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   1440
      TabIndex        =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame frmSetupMenu 
         Caption         =   "Setup Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   360
         TabIndex        =   121
         Top             =   480
         Width           =   3135
         Begin VB.CommandButton cmdCloseSetup 
            Caption         =   "Close"
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
            TabIndex        =   123
            Top             =   3120
            Width           =   2775
         End
         Begin VB.CommandButton cmdSetupOpeningCash 
            Caption         =   "Setup Opening Cash"
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
            TabIndex        =   122
            Top             =   480
            Width           =   2775
         End
      End
      Begin VB.Frame frmOpeningCash 
         Caption         =   "Opening Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   360
         TabIndex        =   124
         Top             =   480
         Visible         =   0   'False
         Width           =   7815
         Begin VB.PictureBox picOpeningCash 
            Height          =   2415
            Left            =   5520
            ScaleHeight     =   2355
            ScaleWidth      =   1755
            TabIndex        =   129
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdOpeningCashClose 
            Caption         =   "Close"
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
            Left            =   2280
            TabIndex        =   128
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CommandButton cmdOpeningCashOK 
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
            Left            =   2280
            TabIndex        =   127
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtOpeningCash 
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
            Left            =   2280
            TabIndex        =   125
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label lblOpeingCash 
            Caption         =   "Enter Opening Cash"
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
            TabIndex        =   126
            Top             =   720
            Width           =   2175
         End
      End
   End
   Begin VB.Frame frmEmployee 
      Caption         =   "Employee File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CommandButton cmdCloseEmployee 
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
         Left            =   240
         TabIndex        =   8
         Top             =   6120
         Width           =   2415
      End
      Begin VB.CommandButton cmdEraseFile 
         Caption         =   "Erase File"
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
         Left            =   240
         TabIndex        =   35
         Top             =   5400
         Width           =   2415
      End
      Begin VB.ListBox listEmployee 
         Height          =   4350
         Left            =   6480
         TabIndex        =   30
         Top             =   720
         Width           =   3135
      End
      Begin VB.Frame frmAdd 
         Caption         =   "Add Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   5775
         Begin VB.CommandButton cmdClearEmpName 
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
            Left            =   3240
            TabIndex        =   83
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddEmployee 
            Caption         =   "Add"
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
            Left            =   1080
            TabIndex        =   13
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtAddName 
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
            Left            =   1080
            TabIndex        =   12
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label lblAddName 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame frmUpdate 
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
         ForeColor       =   &H000000FF&
         Height          =   2535
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   5775
         Begin VB.CommandButton cmdClearUpdateEmp 
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
            Left            =   3720
            TabIndex        =   84
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdUpdateEmployee 
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
            Height          =   615
            Left            =   1560
            TabIndex        =   19
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtUpdateId 
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
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtUpdateName 
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
            Left            =   1560
            TabIndex        =   15
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblUpdateId 
            Caption         =   "ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   18
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblUpdateName 
            Caption         =   "New Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Label lblEmployeeList 
         Alignment       =   2  'Center
         Caption         =   "Employee List"
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
         Left            =   6480
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame frmReview 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid msgOldSales 
         Height          =   7335
         Left            =   120
         TabIndex        =   103
         Top             =   360
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   12938
         _Version        =   393216
         Rows            =   2000
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.CommandButton cmdCloseReprint 
         Caption         =   "CLOSE"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   6960
         Width           =   3135
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
         Left            =   4320
         TabIndex        =   20
         Top             =   6240
         Width           =   3135
      End
      Begin VB.Frame frmOldTran 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Old Transactions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   5775
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         Begin VB.Frame frmViewOldTran 
            Height          =   2175
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   2895
            Begin VB.CommandButton cmdViewOldTran 
               Caption         =   "View Transaction"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   112
               Top             =   1080
               Width           =   2655
            End
            Begin VB.ComboBox comYear 
               Height          =   315
               Left            =   1800
               TabIndex        =   111
               Top             =   480
               Width           =   975
            End
            Begin VB.ComboBox comDay 
               Height          =   315
               Left            =   960
               TabIndex        =   110
               Top             =   480
               Width           =   735
            End
            Begin VB.ComboBox comMonth 
               Height          =   315
               Left            =   120
               TabIndex        =   109
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblYear 
               Caption         =   "Year"
               Height          =   255
               Left            =   1800
               TabIndex        =   108
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblDay 
               Caption         =   "Day"
               Height          =   255
               Left            =   960
               TabIndex        =   107
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblMonth 
               Caption         =   "Month"
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame frmOtherTran 
            Height          =   2295
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
            Width           =   2895
            Begin VB.CommandButton cmdViewOther 
               Caption         =   "View Transaction For Other Day"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   118
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label lblOldTran 
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
               Height          =   1095
               Left            =   120
               TabIndex        =   119
               Top             =   120
               Width           =   2655
            End
         End
         Begin VB.Frame frmPrintOldTran 
            Caption         =   "Print Transaction"
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
            Height          =   2895
            Left            =   120
            TabIndex        =   2
            Top             =   2760
            Visible         =   0   'False
            Width           =   2895
            Begin VB.CommandButton cmdSalesView 
               Caption         =   "View Sales Summery"
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
               TabIndex        =   37
               Top             =   1920
               Width           =   2655
            End
            Begin VB.CommandButton cmdPrintOldTran 
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
               Height          =   735
               Left            =   120
               TabIndex        =   5
               Top             =   1200
               Width           =   2655
            End
            Begin VB.TextBox txtOldTranNum 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   120
               TabIndex        =   3
               Top             =   600
               Width           =   2655
            End
            Begin VB.Label lblOldTranNum 
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
               TabIndex        =   4
               Top             =   360
               Width           =   2655
            End
         End
      End
      Begin VB.Frame frmEmpCommision 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Commision View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   7560
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdPrintOldCommision 
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
            Height          =   495
            Left            =   2040
            TabIndex        =   116
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdClearOldCommision 
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
            Height          =   495
            Left            =   2040
            TabIndex        =   115
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdShowTips 
            Caption         =   "Show Tips"
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
            Left            =   120
            TabIndex        =   113
            Top             =   1320
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid msgOldCommision 
            Height          =   5295
            Left            =   120
            TabIndex        =   104
            Top             =   1920
            Width           =   3940
            _ExtentX        =   6959
            _ExtentY        =   9340
            _Version        =   393216
            Rows            =   50
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
         End
         Begin VB.CommandButton cmdShowCommision 
            Caption         =   "Show Commision"
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
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox comEmpList 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Text            =   "Select Employee"
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame frmViewBack 
         BackColor       =   &H00C0C0FF&
         Height          =   7455
         Left            =   7560
         TabIndex        =   114
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frmMenuFile 
      Caption         =   "Menu File"
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
      Height          =   8535
      Left            =   240
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton cmdEraseMenuFile 
         Caption         =   "Erase File"
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
         Left            =   6480
         TabIndex        =   93
         Top             =   7560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCloseUpdateAdd 
         Caption         =   "Close"
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
         Left            =   8760
         TabIndex        =   78
         Top             =   7560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame frmItemList 
         Caption         =   "Item List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   6240
         TabIndex        =   77
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid msgItemList 
            Height          =   6375
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   5040
            _ExtentX        =   8890
            _ExtentY        =   11245
            _Version        =   393216
            Rows            =   100
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
         End
      End
      Begin VB.Frame frmAddItem 
         Caption         =   "Add Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3615
         Left            =   240
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtCommision 
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
            Left            =   1680
            TabIndex        =   94
            Top             =   1560
            Width           =   3135
         End
         Begin VB.TextBox txtItemPrice 
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
            Left            =   1680
            TabIndex        =   66
            Top             =   960
            Width           =   3135
         End
         Begin VB.CommandButton cmdClearItem 
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
            Left            =   1680
            TabIndex        =   64
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox txtItemName 
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
            Left            =   1680
            TabIndex        =   62
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "Add"
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
            Left            =   1680
            TabIndex        =   61
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblCommision 
            Caption         =   "Commision"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblItemPrice 
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblItemName 
            Caption         =   "New Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame frmUpdateItem 
         Caption         =   "Update Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4335
         Left            =   240
         TabIndex        =   59
         Top             =   4080
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtNewCommision 
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
            Left            =   2040
            TabIndex        =   96
            Top             =   2280
            Width           =   3015
         End
         Begin VB.CommandButton cmdClearUpdateItem 
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
            Left            =   2040
            TabIndex        =   90
            Top             =   3480
            Width           =   2175
         End
         Begin VB.CommandButton cmdUpdateItem 
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
            Height          =   615
            Left            =   2040
            TabIndex        =   89
            Top             =   2880
            Width           =   2175
         End
         Begin VB.TextBox txtNewItemPrice 
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
            Left            =   2040
            TabIndex        =   88
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txtNewItemName 
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
            Left            =   2040
            TabIndex        =   87
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtOldId 
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
            Left            =   2040
            TabIndex        =   85
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblNewCommision 
            Caption         =   "New/Old Commision"
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
            Left            =   240
            TabIndex        =   97
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblNewPrice 
            Caption         =   "New/Old Price"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblNewName 
            Caption         =   "New/Old Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblOldId 
            Caption         =   "Old Item  ID "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame frmMenuUpdateAdd 
         Height          =   6255
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   4935
         Begin VB.CommandButton cmdOthers 
            Caption         =   "Others Services"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2400
            TabIndex        =   80
            Top             =   3720
            Width           =   2175
         End
         Begin VB.CommandButton cmdProduct 
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
            Height          =   855
            Left            =   240
            TabIndex        =   79
            Top             =   3720
            Width           =   2175
         End
         Begin VB.CommandButton cmdDiscount 
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2400
            TabIndex        =   76
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton cmdMassage 
            Caption         =   "Massage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2400
            TabIndex        =   73
            Top             =   2040
            Width           =   2175
         End
         Begin VB.CommandButton cmdMenuFileClose 
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
            Height          =   855
            Left            =   240
            TabIndex        =   75
            Top             =   5160
            Width           =   2175
         End
         Begin VB.CommandButton cmdSpecial 
            Caption         =   "Special"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            TabIndex        =   74
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton cmdThreading 
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
            Height          =   855
            Left            =   2400
            TabIndex        =   71
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdTips 
            Caption         =   "Tips && Wraps"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2400
            TabIndex        =   70
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdWaxing 
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
            Height          =   855
            Left            =   240
            TabIndex        =   72
            Top             =   2040
            Width           =   2175
         End
         Begin VB.CommandButton cmdPedicure 
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
            Height          =   855
            Left            =   240
            TabIndex        =   69
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdManicure 
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
            Height          =   855
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   6480
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6480
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "Program"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "POS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rNum As Integer
Dim menuFile As String
Dim updateId As Integer
Dim prefix As String
Dim iName As String
Dim maxItemNum As Integer
Dim itemNum As Integer

Private Sub cmdAddEmployee_Click()
    Dim emp As employee
    
    If Trim(txtAddName.Text) = "" Or Trim(txtAddName.Text) = "Enter the name" Or Trim(txtAddName.Text) = "Name already used" Then
        Let txtAddName.Text = "Enter the name"
        Exit Sub
    Else
        If checkEmpName(txtAddName.Text) Then
            Let emp.name = UCase(txtAddName.Text)
        Else
            Let txtAddName.Text = "Name already used"
            txtAddName.SetFocus
            Exit Sub
        End If
    End If
        
    Let empNum = empNum + 1
    Let emp.id = empNum
    Put #2, empNum, emp
    listEmployee.Clear
    listEmployee.AddItem "ID     Employee Name"
    listEmployee.AddItem "----     --------------------------------------------------"
    For i = 1 To empNum
        Get #2, i, emp
        listEmployee.AddItem emp.id & "       " & emp.name
    Next i
End Sub

Private Sub cmdAddItem_Click()
    Dim item As product
    
    If itemNum >= maxItemNum Then
        Let info = "More than " & maxItemNum & " items cannot be created for '" & iName & "'. Update unused item instead."
        MsgBox (info)
        Exit Sub
    End If
    If Trim(txtItemName.Text) = "" Or Trim(txtItemName.Text) = "Enter the name" Or Trim(txtItemName.Text) = "Name already used" Then
        Let txtItemName.Text = "Enter the name"
        Exit Sub
    Else
        If checkProductName(Trim(txtItemName.Text)) Then
            Let item.name = UCase(txtItemName.Text)
        Else
            Let txtItemName.Text = "Name already used"
            txtItemName.SetFocus
            Exit Sub
        End If
    End If
    If IsNumeric(Trim(txtItemPrice.Text)) Then
        Let item.price = Val(Trim(txtItemPrice.Text))
    Else
        Let txtItemPrice.Text = "Enter Price"
        txtItemPrice.SetFocus
        Exit Sub
    End If
    If IsNumeric(Trim(txtCommision.Text)) Then
        Let item.commision = Val(Trim(txtCommision.Text))
    Else
        Let item.commision = 0
    End If
    
    Let itemNum = itemNum + 1
    Let item.productId = prefix & itemNum
    
    Put #2, itemNum, item
    
    Call PrintItemList(itemNum)
    Let txtItemName.Text = ""
    Let txtItemPrice.Text = ""
    Let txtCommision.Text = "0"
    txtItemName.SetFocus
End Sub

Private Sub cmdClearEmpName_Click()
    Let txtAddName.Text = ""
End Sub

Private Sub cmdClearOldCommision_Click()
    msgOldCommision.Clear
End Sub

Private Sub cmdClearTran_Click()
    Call ClearMsg(msgOldSales)
End Sub

Private Sub cmdClearUpdateEmp_Click()
    Let txtUpdateId.Text = ""
    Let txtUpdateName.Text = ""
End Sub

Private Sub cmdClearUpdateItem_Click()
    Let txtOldId.Text = ""
    Let txtNewItemName.Text = ""
    Let txtNewItemPrice.Text = ""
    Let txtNewCommision.Text = ""
    
End Sub


Private Sub cmdCloseEmployee_Click()
    frmDisplay.Visible = True
    frmEmployee.Visible = False
    
    Close #2
End Sub

Private Sub cmdCloseReprint_Click()
    Call ClearMsg(msgOldSales)
    frmPrintOldTran.Visible = False
    frmEmpCommision.Visible = False
    frmViewOldTran.Visible = True
    frmOtherTran.Visible = False
    frmViewBack.Visible = True
    frmReview.Visible = False
    frmDisplay.Visible = True
    
End Sub

Private Sub cmdCloseSetup_Click()
    frmDisplay.Visible = True
    frmSetup.Visible = False
End Sub

Private Sub cmdCloseUpdateAdd_Click()
    msgItemList.Clear
    frmMenuUpdateAdd.Visible = True
    frmAddItem.Visible = False
    frmUpdateItem.Visible = False
    frmItemList.Visible = False
    cmdCloseUpdateAdd.Visible = False
    cmdEraseMenuFile.Visible = False
    Let frmMenuFile.Caption = "Menu File"
    
    Close #2
End Sub

Private Sub cmdDiscount_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Discount"
    
    Dim item As product
    Let prefix = "DIS"
    Let iName = "DISCOUNT"
    Let menuFile = defaultDir & "\nailsPOS\menu\discount.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = 200
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdEmployeeFile_Click()
    Dim emp As employee
    
    frmEmployee.Visible = True
    frmDisplay.Visible = False
     
    
    Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #2 Len = Len(emp)
    Let empNum = LOF(2) / Len(emp)
    
    listEmployee.Clear
    listEmployee.AddItem "ID     Employee Name"
    listEmployee.AddItem "----     --------------------------------------------------"
    For i = 1 To empNum
        Get #2, i, emp
        listEmployee.AddItem emp.id & "       " & emp.name
    Next i
    
End Sub

Private Sub cmdEraseFile_Click()
    Dim emp As employee
    Let yn = MsgBox("Are you sure you want to erase Employee File?", vbYesNo)
    If yn = 6 Then
        Close #2
        Kill defaultDir & "\nailsPOS\setup\employee.txt"
        Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #2 Len = Len(emp)
        Let empNum = LOF(2) / Len(emp)
    
        listEmployee.Clear
        listEmployee.AddItem "ID     Employee Name"
        listEmployee.AddItem "----     --------------------------------------------------"
        For i = 1 To empNum
            Get #2, i, emp
            listEmployee.AddItem emp.id & "       " & emp.name
        Next i
    End If
End Sub

Private Sub cmdEraseMenuFile_Click()
    Dim item As product
    Let yn = MsgBox("Are you sure you want to erase this Menu File?", vbYesNo)
    If yn = 6 Then
        
    End If
End Sub

Private Sub cmdLogin_Click()
    If (UCase(txtPassword.Text) = "MGR") Then
        cmdLogin.Visible = False
        cmdLogout.Visible = True
        cmdRegister.Visible = False
        txtPassword.Visible = False
        txtPassword.Text = ""
        lblPassword.Visible = False
        frmManagerTask.Visible = True
        frmManagerTaskBk.Visible = False
        mnuProgram.Visible = False
    Else
        Let txtPassword.PasswordChar = ""
        Let txtPassword.Text = "Wrong. Try Again."
    End If
End Sub

Private Sub cmdLogout_Click()
    cmdRegister.Visible = True
    txtPassword.Visible = True
    lblPassword.Visible = True
    cmdLogout.Visible = False
    cmdLogin.Visible = True
    frmManagerTask.Visible = False
    frmManagerTaskBk.Visible = True
    mnuProgram.Visible = True
End Sub

Private Sub cmdManicure_Click()
        
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Manicure"
    
    Dim item As product
    Let prefix = "MAN"
    Let iName = "MANICURE"
    Let menuFile = defaultDir & "\nailsPOS\menu\manicure.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdManicure.Count
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
    
End Sub

Private Sub cmdMassage_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Massage"
    
    Dim item As product
    Let prefix = "MAS"
    Let iName = "MASSAGE"
    Let menuFile = defaultDir & "\nailsPOS\menu\massage.txt"
    
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdMassage.Count
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdMenuFile_Click()
    frmMenuFile.Visible = True
    frmDisplay.Visible = False
End Sub

Private Sub cmdMenuFileClose_Click()
    frmMenuFile.Visible = False
    frmDisplay.Visible = True
End Sub

Private Sub cmdOldIdOK_Click()
    Dim item As product
    If Trim(txtOldId.Text) = "" Or Trim(txtOldId.Text) = "Enter Product ID" Or Trim(txtOldId.Text) = "ID not found" Then
        Let txtItemName.Text = "Enter Product ID"
        Exit Sub
    Else
        Let updateId = checkProductID(Trim(txtOldId.Text))
        If updateId = 0 Then
            Let txtOldId.Text = "ID not found"
            txtItemName.SetFocus
            Exit Sub
        Else
            Let item.productId = UCase(txtOldId.Text)
            frmUpdateItem1.Visible = True
            frmOldId.Visible = False
        End If
    End If
End Sub

Private Sub cmdOpeningCashClose_Click()
    Let txtOpeningCash.Text = ""
    picOpeningCash.Cls
    frmOpeningCash.Visible = False
    frmSetupMenu.Visible = True
End Sub

Private Sub cmdOpeningCashOK_Click()
    Dim startclosingcash As openingCash
    Dim newCash As Single
        
    If IsNumeric(Trim(txtOpeningCash.Text)) Then
        Let newCash = Val(Trim(txtOpeningCash.Text))
    Else
        Let txtOpeningCash.Text = "Enter Again."
        Exit Sub
    End If
    
    Open defaultDir & "\nailsPOS\setup\openingCash.txt" For Random As #3 Len = Len(startclosingcash)
    
    Let startclosingcash.registerCash = newCash
    Put #3, 1, startclosingcash
    Close #3
    picOpeningCash.Cls
    picOpeningCash.Print "New Opening Cash"
    picOpeningCash.Print "---------------------"
    picOpeningCash.Print Format(newCash, "currency")
    Let txtOpeningCash.Text = ""
    
End Sub

Private Sub cmdOthers_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Other Services"
    
    Dim item As product
    Let prefix = "OTH"
    Let menuFile = defaultDir & "\nailsPOS\menu\others.txt"
    Let iName = "OTHER SERVICES"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdOthers.Count
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdPedicure_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Pedicure"
    
    Dim item As product
    Let prefix = "PED"
    Let iName = "PEDICURE"
    Let menuFile = defaultDir & "\nailsPOS\menu\pedicure.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdPedicure.Count
    Call PrintItemList(itemNum)
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdPrintOldCommision_Click()
    Dim s As String
    Let msgOldCommision.row = 0
    Let msgOldCommision.Col = 0
    Printer.Print msgOldCommision.Text
    
    Let msgOldCommision.row = 1
    Let msgOldCommision.Col = 0
    Let p = msgOldCommision.Text
    Let msgOldCommision.Col = 1
    Let p = p & msgOldCommision.Text
    Printer.Print p
    
    For i = 2 To 30
        Let msgOldCommision.row = i
        Let msgOldCommision.Col = 0
        Let s = Trim(msgOldCommision.Text)
        If s = "" Then
            Exit For
        End If
        Let msgOldCommision.Col = 1
        Call PrintReceipt(s, Val(Trim(msgOldCommision.Text)))
        
    Next i
    Printer.Print
    Printer.Print
    Printer.Print storeName
    Printer.EndDoc
End Sub

Private Sub cmdPrintOldTran_Click()
    Dim tranNum As Integer
    Dim item As salesItem
    
    If IsNumeric(txtOldTranNum.Text) = False Then
        Let txtOldTranNum.Text = "Not Valid Number"
        Exit Sub
    End If
    
    Let tranNum = Val(Trim(txtOldTranNum.Text))
        
    Open oldFileName For Random As #3 Len = Len(item)
        
    Call ResetSales
        
    For record = 1 To rNum
        Get #3, record, item
        If tranNum = item.tranNum Then
            If IsNumeric(Trim(item.itemType)) = True Then
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
                    Let tstr = item.timeStr
                    Let dstr = item.dateStr
                    Let balance = item.price
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
    Close #3
        
    Call PrintHead(dstr, tstr, tranNum)
    
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
    
    
End Sub

Private Sub cmdProduct_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Product"
    
    Dim item As product
    Let prefix = "PRO"
    Let iName = "PRODUCT"
    Let menuFile = defaultDir & "\nailsPOS\menu\product.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = 200
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdProfileClose_Click()
    frmStoreProfile.Visible = False
    frmDisplay.Visible = True
    Close #2
End Sub

Private Sub cmdProfileUpdate_Click()
    Dim storeInfo As profile
    
    Let storeInfo.name = Trim(txtStoreName.Text)
    Let storeInfo.street = Trim(txtStreet.Text)
    Let storeInfo.city = Trim(txtCity.Text)
    Let storeInfo.state = Trim(txtState.Text)
    Let storeInfo.zip = Val(Trim(txtZip.Text))
    Let storeInfo.phone = Trim(txtPhone.Text)
    Let storeInfo.web = Trim(txtWeb.Text)
    
    Put #2, 1, storeInfo
    
    picProfileView.Cls
    Get #2, 1, storeInfo
    picProfileView.Print Trim(storeInfo.name)
    picProfileView.Print Trim(storeInfo.street)
    picProfileView.Print Trim(storeInfo.city) & ", " & Trim(storeInfo.state) & ", " & Trim(storeInfo.zip)
    picProfileView.Print Trim(storeInfo.phone)
    picProfileView.Print Trim(storeInfo.web)
End Sub

Private Sub cmdRegister_Click()
           
    If Dir(defaultDir & "\nailsPOS\sales\closed" & fileName) <> "" Then
        MsgBox ("Today's Business is already closed. Can't open new transaction for today.")
        Exit Sub
    Else
        mnuProgram.Visible = False
        POS.Hide
        Register.Show
        Register.cmdManicure(0).SetFocus
    End If
        
End Sub

Private Sub cmdClearItem_Click()
    Let txtItemName.Text = ""
    Let txtItemPrice.Text = ""
    Let txtCommision.Text = ""
    txtItemName.SetFocus
End Sub

Private Sub cmdReview_Click()
    frmReview.Visible = True
    frmDisplay.Visible = False
    
    Let msgOldCommision.ColWidth(0) = 2700
    Let msgOldCommision.ColWidth(1) = 900
    Let msgOldCommision.ColAlignment(0) = 1
    
    Let msgOldSales.ColWidth(0) = 600
    Let msgOldSales.ColWidth(1) = 2400
    Let msgOldSales.ColWidth(2) = 800
    Let msgOldSales.ColAlignment(1) = 1
    Let msgOldSales.ColAlignment(0) = 3
    Let msgOldSales.row = 0
    Let msgOldSales.Col = 0
    Let msgOldSales.CellFontUnderline = True
    Let msgOldSales.CellFontBold = True
    Let msgOldSales.Text = Format("T. No", "@@@@@")
    Let msgOldSales.Col = 1
    Let msgOldSales.CellFontUnderline = True
    Let msgOldSales.CellFontBold = True
    Let msgOldSales.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgOldSales.Col = 2
    Let msgOldSales.CellFontUnderline = True
    Let msgOldSales.CellFontBold = True
    Let msgOldSales.Text = Format("Price", "@@@@@@@@@")
    
    
    Dim emp As employee
    
    Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #3 Len = Len(emp)
    Let empNum = LOF(3) / Len(emp)
    
    For i = 1 To empNum
        Get #3, i, emp
        comEmpList.AddItem emp.id & "  " & emp.name
    Next i
    Close #3
End Sub

Private Sub cmdSalesView_Click()
    
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
    
   
    Open oldFileName For Random As #3 Len = Len(item)
    
    For record = 1 To rNum
        Get #3, record, item
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
            Let dstr = item.dateStr
        End If
    Next record
    
    msgOldCommision.Clear
    msgOldCommision.row = 0
    msgOldCommision.Col = 0
    msgOldCommision.CellFontBold = True
    msgOldCommision.Text = "Sales Summary: " & dstr
    msgOldCommision.row = 1
    msgOldCommision.Col = 0
    msgOldCommision.Text = "----------------------------------------------------"
    msgOldCommision.Col = 1
    msgOldCommision.Text = "--------------------------------------"
    msgOldCommision.row = 2
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total Gross Sales: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(subTotal, "0.00")
    msgOldCommision.row = 3
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total GiftCertificate sold: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tGC, "0.00")
    msgOldCommision.row = 4
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total Sales Tax: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tTax, "0.00")
    msgOldCommision.row = 5
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total GiftCertificate Redemed: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tGCR, "0.00")
    msgOldCommision.row = 6
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total Commision: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tCommision, "0.00")
    msgOldCommision.row = 7
    msgOldCommision.Col = 0
    msgOldCommision.CellFontBold = True
    msgOldCommision.Text = "Total Amount in Register: "
    msgOldCommision.Col = 1
    Let net = tSales - tGCR - tCommision
    msgOldCommision.Text = Format(net, "0.00")
    msgOldCommision.row = 8
    msgOldCommision.Col = 0
    msgOldCommision.CellFontBold = True
    msgOldCommision.Text = "Total Cash in register: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(net - tCC - tTips, "0.00")
    msgOldCommision.row = 9
    msgOldCommision.Col = 0
    msgOldCommision.CellFontBold = True
    msgOldCommision.Text = "Total Credit Card: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tCC + tTips, "0.00")
    msgOldCommision.row = 10
    msgOldCommision.Col = 0
    msgOldCommision.Text = "Total Credit Card tips: "
    msgOldCommision.Col = 1
    msgOldCommision.Text = Format(tTips, "0.00")
    Close #3
End Sub

Private Sub cmdSetup_Click()
    frmSetup.Visible = True
    frmDisplay.Visible = False
End Sub

Private Sub cmdSetupOpeningCash_Click()
    Dim startclosingcash As openingCash
    frmSetupMenu.Visible = False
    frmOpeningCash.Visible = True
    
    Open defaultDir & "\nailsPOS\setup\openingCash.txt" For Random As #3 Len = Len(startclosingcash)
    Get #3, 1, startclosingcash
    
    picOpeningCash.Cls
    picOpeningCash.Print "Opening Cash"
    picOpeningCash.Print "---------------------"
    picOpeningCash.Print Format(startclosingcash.registerCash, "currency")
    
    Close #3
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
    
    msgOldCommision.Clear
    Let id = comEmpList.ListIndex + 1
    
    Let msgOldCommision.row = 0
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontBold = True
    
    If id = 0 Then
        Let msgOldCommision.Text = "Select Employee First"
        Exit Sub
    End If
    Let msgOldCommision.Text = comEmpList.Text
    
    Open oldFileName For Random As #3 Len = Len(item)
       
    Let msgOldCommision.row = 1
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontUnderline = True
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgOldCommision.Col = 1
    Let msgOldCommision.CellFontUnderline = True
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format("Commision", "@@@@@@@@@")
    
    For r = 1 To rNum
        Get #3, r, item
        If Trim(item.itemType) = 100 And id = item.price Then
            Let match = True
        Else
            If Trim(item.itemType) = 102 Then
                Let match = False
            ElseIf (match) Then
                If item.commision > 0 Then
                    Let row = row + 1
                    Let msgOldCommision.row = row
                    Let msgOldCommision.Col = 0
                    Let msgOldCommision.Text = Trim(item.name)
                    Let msgOldCommision.Col = 1
                    Let msgOldCommision.Text = Format(item.commision, "0.00")
                    Let commision = commision + item.commision
                End If
            End If
        End If
    Next r
    Let msgOldCommision.row = row + 1
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = "Total"
    Let msgOldCommision.Col = 1
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format(commision, "0.00")
    Close #3
End Sub

Private Sub cmdShowTips_Click()
    Dim item As salesItem
    Dim r As Integer
    Dim tips As Single
    Dim id As Integer
    Dim row As Integer
    Dim match As Boolean
    
    Let match = False
    Let row = 1
    Let tips = 0
    
    msgOldCommision.Clear
    Let id = comEmpList.ListIndex + 1
    
    Let msgOldCommision.row = 0
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontBold = True
    
    If id = 0 Then
        Let msgOldCommision.Text = "Select Employee First"
        Exit Sub
    End If
    
    Let msgOldCommision.Text = comEmpList.Text
    Open oldFileName For Random As #3 Len = Len(item)
     
    
    Let msgOldCommision.row = 1
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontUnderline = True
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format("Credit Card Tips", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgOldCommision.Col = 1
    Let msgOldCommision.CellFontUnderline = True
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format("Amount", "@@@@@@@@@")
    
    For r = 1 To rNum
        Get #3, r, item
        If Trim(item.itemType) = 100 Then
            If id = Trim(item.price) Then
                Let match = True
            End If
            Let r = r + 6
        Else
            If Trim(item.itemType) = 109 Then
                If (match) Then
                    If item.price > 0 Then
                        Let row = row + 1
                        Let msgOldCommision.row = row
                        Let msgOldCommision.Col = 0
                        Let msgOldCommision.Text = "Transaction #" & Trim(item.tranNum)
                        Let msgOldCommision.Col = 1
                        Let msgOldCommision.Text = Format(item.price, "0.00")
                        Let tips = tips + item.price
                    End If
                End If
                Let match = False
            End If
        End If
    Next r
    Let msgOldCommision.row = row + 1
    Let msgOldCommision.Col = 0
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = "Total"
    Let msgOldCommision.Col = 1
    Let msgOldCommision.CellFontBold = True
    Let msgOldCommision.Text = Format(tips, "0.00")
    Close #3
End Sub

Private Sub cmdSpecial_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Special"
    
    Dim item As product
    Let prefix = "SPE"
    Let iName = "SPECIAL"
    Let menuFile = defaultDir & "\nailsPOS\menu\special.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdSpecial.Count
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdStoreProfile_Click()
    frmStoreProfile.Visible = True
    frmDisplay.Visible = False
    
    Dim storeInfo As profile
    Open defaultDir & "\nailsPOS\setup\profile.txt" For Random As #2 Len = Len(storeInfo)
    
    picProfileView.Cls
    Get #2, 1, storeInfo
    picProfileView.Print Trim(storeInfo.name)
    picProfileView.Print Trim(storeInfo.street)
    picProfileView.Print Trim(storeInfo.city) & ", " & Trim(storeInfo.state) & ", " & Trim(storeInfo.zip)
    picProfileView.Print Trim(storeInfo.phone)
    picProfileView.Print Trim(storeInfo.web)
   
End Sub

Private Sub cmdThreading_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Threading"
    
    Dim item As product
    Let prefix = "THR"
    Let iName = "THREADING"
    Let menuFile = defaultDir & "\nailsPOS\menu\threading.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdThreading.Count
    Call PrintItemList(itemNum)
    
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdTips_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Tips && Warps"
    
    Dim item As product
    Let prefix = "TIP"
    Let iName = "TIPS AND WRAPS"
    Let menuFile = defaultDir & "\nailsPOS\menu\tips.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdTips.Count
    Call PrintItemList(itemNum)
    cmdUpdateItem.SetFocus
End Sub

Private Sub cmdUpdateEmployee_Click()
    Dim emp As employee
    
    If Trim(txtUpdateName.Text) = "" Or Trim(txtUpdateName.Text) = "Enter the name" Or Trim(txtUpdateName.Text) = "Name already used" Then
        Let txtUpdateName.Text = "Enter the name"
        Exit Sub
    Else
        If checkEmpName(Trim(txtUpdateName.Text)) Then
            Let emp.name = UCase(Trim(txtUpdateName.Text))
        Else
            Let txtUpdateName.Text = "Name already used"
            txtUpdateName.SetFocus
            Exit Sub
        End If
    End If
    If IsNumeric(Trim(txtUpdateId.Text)) Then
        Let temp = Val(Trim(txtUpdateId.Text))
        If (temp > 0 And temp <= empNum) Then
            Let emp.id = temp
        Else
            Let txtUpdateId.Text = "ID not found"
            txtUpdateId.SetFocus
            Exit Sub
        End If
    Else
        Let txtUpdateId.Text = "Enter a Number"
        txtUpdateId.SetFocus
        Exit Sub
    End If
    
    Put #2, emp.id, emp
    listEmployee.Clear
    listEmployee.AddItem "ID     Employee Name"
    listEmployee.AddItem "----     --------------------------------------------------"
    For i = 1 To empNum
        Get #2, i, emp
        listEmployee.AddItem emp.id & "       " & emp.name
    Next i
End Sub

Private Sub cmdUpdateItem_Click()
    Dim item As product
    
    If Trim(txtOldId.Text) = "" Or Trim(txtOldId.Text) = "Enter Item ID" Or Trim(txtOldId.Text) = "ID not found" Then
        Let txtOldId.Text = "Enter Item ID"
        txtOldId.SetFocus
        Exit Sub
    Else
        Let updateId = checkProductID(Trim(txtOldId.Text))
        If updateId = 0 Then
            Let txtOldId.Text = "ID not found"
            txtOldId.SetFocus
            Exit Sub
        Else
            Let item.productId = prefix & updateId
        End If
    End If
    
    If Trim(txtNewItemName.Text) = "" Then
        Let txtNewItemName.Text = "Enter name"
        txtNewItemName.SetFocus
        Exit Sub
    Else
        Let item.name = UCase(Trim(txtNewItemName.Text))
    End If
    
    If IsNumeric(Trim(txtNewItemPrice.Text)) Then
        Let item.price = Val(Trim(txtNewItemPrice.Text))
    Else
        Let txtNewItemPrice.Text = "Enter Price"
        txtNewItemPrice.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Trim(txtNewCommision.Text)) Then
        Let item.commision = Val(Trim(txtNewCommision.Text))
    Else
        Let txtNewCommision.Text = "0"
        txtNewCommision.SetFocus
        Exit Sub
    End If
    
           
    Put #2, updateId, item
    
    Call PrintItemList(itemNum)
    Let txtOldId.Text = ""
    Let txtNewItemName.Text = ""
    Let txtNewItemPrice.Text = ""
    Let txtCommision.Text = ""
    
End Sub

Private Sub cmdViewOldTran_Click()
    Dim item As salesItem
    Dim tran As Integer
    
    Let oldFileName = defaultDir & "\nailsPOS\sales\closedsale" & comMonth.Text & comDay.Text & comYear.Text & ".txt"
    
    
    If Dir(oldFileName) <> "" Then
        Call ClearMsg(msgOldSales)
        Open oldFileName For Random As #3 Len = Len(item)
        Let rNum = LOF(3) / Len(item)
                
        For record = 1 To rNum
            Get #3, record, item
            If tran < item.tranNum Then
                msgOldSales.row = record
                msgOldSales.Col = 0
                msgOldSales.Text = item.tranNum
                msgOldSales.Col = 1
                msgOldSales.Text = Trim(item.name)
                msgOldSales.Col = 2
                msgOldSales.Text = Format(item.price, "0.00")
                tran = item.tranNum
            Else
                msgOldSales.row = record
                msgOldSales.Col = 1
                msgOldSales.Text = Trim(item.name)
                msgOldSales.Col = 2
                msgOldSales.Text = Format(item.price, "0.00")
            End If
        Next record
        Close #3
        frmPrintOldTran.Visible = True
        frmEmpCommision.Visible = True
        frmViewOldTran.Visible = False
        frmOtherTran.Visible = True
        frmViewBack.Visible = False
        Let lblOldTran.Caption = "This is Transactions of " & comMonth.Text & "-" & comDay.Text & "-" & comYear.Text
        
    Else
        msgOldSales.row = 1
        msgOldSales.Col = 1
        msgOldSales.Text = "File not found."
    End If
    
    Let comMonth.Text = ""
    Let comDay.Text = ""
    Let comYear.Text = ""
End Sub

Private Sub cmdViewOther_Click()
    Call ClearMsg(msgOldSales)
    msgOldCommision.Clear
    frmPrintOldTran.Visible = False
    frmEmpCommision.Visible = False
    frmViewOldTran.Visible = True
    frmOtherTran.Visible = False
    frmViewBack.Visible = True
    
End Sub

Private Sub cmdWaxing_Click()
    frmMenuUpdateAdd.Visible = False
    frmAddItem.Visible = True
    frmUpdateItem.Visible = True
    frmItemList.Visible = True
    cmdCloseUpdateAdd.Visible = True
    cmdEraseMenuFile.Visible = True
    Let frmMenuFile.Caption = "Waxing"
    
    Dim item As product
    Let prefix = "WAX"
    Let iName = "WAXING"
    Let menuFile = defaultDir & "\nailsPOS\menu\waxing.txt"
    Open menuFile For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    Let maxItemNum = Register.cmdWaxing.Count
    Call PrintItemList(itemNum)
    cmdUpdateItem.SetFocus
End Sub



Private Sub Form_Load()
    Let lblToday.Caption = "Today is: " & Format(Date, "dddd, mmmm dd, yyyy")
    Let defaultDir = CurDir()
    Call AddMonth
    Call AddDay
    Call AddYear
    Call SetManicure
    Call SetPedicure
    Call SetWaxing
    Call SetTips
    Call SetThreading
    Call SetMassage
    Call SetProduct
    Call SetSpecial
    Call SetOthers
    Call SetDiscount
    Call SetProfile
    POS.Caption = storeName
    
    
End Sub

Private Sub frmUpdateItem1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub msgItemList_DblClick()
    Let msgItemList.Col = 0
    Let txtOldId.Text = msgItemList.Text
    Let msgItemList.Col = 1
    Let txtNewItemName.Text = msgItemList.Text
    Let msgItemList.Col = 2
    Let txtNewItemPrice.Text = msgItemList.Text
    Let msgItemList.Col = 3
    Let txtNewCommision.Text = msgItemList.Text
End Sub

Private Sub tmrCurrentTime_Timer()
    Let lblTime.Caption = Format(Time, "h:nn:ss AM/PM")
End Sub

Public Sub AddMonth()
    Dim i As Integer
    For i = 1 To 12
        comMonth.AddItem Format(i, "00")
    Next i
End Sub

Public Sub AddDay()
    Dim i As Integer
    For i = 1 To 31
        comDay.AddItem Format(i, "00")
    Next i
End Sub

Public Sub AddYear()
    Dim i As Integer
    For i = 2008 To 2020
        comYear.AddItem i
    Next i
End Sub
Public Function checkProductName(name As String) As Boolean
    Dim item As product
    For i = 1 To itemNum
        Get #2, i, item
        If RTrim(item.name) = UCase(name) Then
            checkProductName = False
            Exit Function
        End If
    Next i
    checkProductName = True
End Function

Public Function checkProductID(id As String) As Integer
    Dim item As product
    For i = 1 To itemNum
        Get #2, i, item
        If RTrim(item.productId) = UCase(id) Then
            checkProductID = i
            Exit Function
        End If
    Next i
    checkProductID = 0
End Function

Public Sub SetManicure()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\manicure.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let manicureName(i) = StrConv(Trim(item.name), 3)
        Let manicurePrice(i) = item.price
        Let Register.cmdManicure(i - 1).Caption = manicureName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 12
        Let manicureName(j) = ""
        Let manicurePrice(j) = 0
        Let Register.cmdManicure(j - 1).Caption = ""
    Next j
    Close #2
End Sub

Public Sub SetPedicure()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\pedicure.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let pedicureName(i) = StrConv(Trim(item.name), 3)
        Let pedicurePrice(i) = item.price
        Let pedicureCommision(i) = item.commision
        Let Register.cmdPedicure(i - 1).Caption = pedicureName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 10
        Let pedicureName(j) = ""
        Let pedicurePrice(j) = 0
        Let pedicureCommision(j) = 0
        Let Register.cmdPedicure(j - 1).Caption = ""
    Next j
    Close #2
End Sub
Public Sub SetWaxing()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\waxing.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let waxingName(i) = StrConv(Trim(item.name), 3)
        Let waxingPrice(i) = item.price
        Let Register.cmdWaxing(i - 1).Caption = waxingName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 14
        Let waxingName(j) = ""
        Let waxingPrice(j) = 0
        Let Register.cmdWaxing(j - 1).Caption = ""
    Next j
    Close #2
End Sub

Public Sub SetTips()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\tips.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let tipsName(i) = StrConv(Trim(item.name), 3)
        Let tipsPrice(i) = item.price
        Let Register.cmdTips(i - 1).Caption = tipsName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 10
        Let tipsName(j) = ""
        Let tipsPrice(j) = 0
        Let Register.cmdTips(j - 1).Caption = ""
    Next j
    Close #2
End Sub
Public Sub SetThreading()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\threading.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let threadingName(i) = StrConv(Trim(item.name), 3)
        Let threadingPrice(i) = item.price
        Let Register.cmdThreading(i - 1).Caption = threadingName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 8
        Let threadingName(j) = ""
        Let threadingPrice(j) = 0
        Let Register.cmdThreading(j - 1).Caption = ""
    Next j
    Close #2
End Sub
Public Sub SetMassage()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\massage.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let massageName(i) = StrConv(Trim(item.name), 3)
        Let massagePrice(i) = item.price
        Let massageCommision(i) = item.commision
        Let Register.cmdMassage(i - 1).Caption = massageName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 6
        Let massageName(j) = ""
        Let massagePrice(j) = 0
        Let massageCommision(j) = 0
        Let Register.cmdMassage(j - 1).Caption = ""
    Next j
    Close #2
End Sub
Public Sub SetSpecial()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\special.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let specialName(i) = StrConv(Trim(item.name), 3)
        Let specialPrice(i) = item.price
        Let specialCommision(i) = item.commision
        Let Register.cmdSpecial(i - 1).Caption = specialName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 6
        Let specialName(j) = ""
        Let specialPrice(j) = 0
        Let specialCommision(j) = 0
        Let Register.cmdSpecial(j - 1).Caption = ""
    Next j
    Close #2
End Sub
Public Sub SetOthers()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\others.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let othersName(i) = StrConv(Trim(item.name), 3)
        Let othersPrice(i) = item.price
        Let othersCommision(i) = item.commision
        Let Register.cmdOthers(i - 1).Caption = othersName(i) & Chr(13) & "$" & item.price
    Next i
    For j = itemNum + 1 To 8
        Let othersName(j) = ""
        Let othersPrice(j) = 0
        Let othersCommision(j) = 0
        Let Register.cmdOthers(j - 1).Caption = ""
    Next j
    Close #2
End Sub

Public Sub SetDiscount()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\discount.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let discountName(i) = StrConv(Trim(item.name), 3)
        Let discountPrice(i) = item.price
        Register.comDiscount.AddItem discountName(i) & " $" & item.price
    Next i
    Close #2
End Sub
Public Sub SetProduct()
    Dim item As product
    Dim f As String
    Let f = defaultDir & "\nailsPOS\menu\product.txt"
    Open f For Random As #2 Len = Len(item)
    Let itemNum = LOF(2) / Len(item)
    For i = 1 To itemNum
        Get #2, i, item
        Let productName(i) = StrConv(Trim(item.name), 3)
        Let productPrice(i) = item.price
        Register.comProductList.AddItem productName(i) & " $" & item.price
    Next i
    Close #2
End Sub

Public Sub PrintItemList(itemNum As Integer)
    Let msgItemList.ColWidth(0) = 800
    Let msgItemList.ColWidth(1) = 2100
    Let msgItemList.ColWidth(2) = 800
    Let msgItemList.ColWidth(3) = 1000
    Let msgItemList.ColAlignment(0) = 4
    Let msgItemList.row = 0
    Let msgItemList.Col = 0
    Let msgItemList.CellFontUnderline = True
    Let msgItemList.CellFontBold = True
    Let msgItemList.Text = Format("Item No", "@@@@@@@")
    Let msgItemList.Col = 1
    Let msgItemList.CellFontUnderline = True
    Let msgItemList.CellFontBold = True
    Let msgItemList.Text = Format("Item Description", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    Let msgItemList.Col = 2
    Let msgItemList.CellFontUnderline = True
    Let msgItemList.CellFontBold = True
    Let msgItemList.Text = Format("Price", "@@@@@@@@@")
    Let msgItemList.Col = 3
    Let msgItemList.CellFontUnderline = True
    Let msgItemList.CellFontBold = True
    Let msgItemList.Text = Format("Commision", "@@@@@@@@@")
    
    Dim p As product
    For i = 1 To itemNum
        Get #2, i, p
        msgItemList.row = i
        msgItemList.Col = 0
        msgItemList.Text = p.productId
        msgItemList.Col = 1
        msgItemList.Text = p.name
        msgItemList.Col = 2
        msgItemList.Text = Format(p.price, "0.00")
        msgItemList.Col = 3
        msgItemList.Text = Format(p.commision, "0.00")
    Next i
End Sub


Public Sub SetProfile()
    Dim storeInfo As profile
    Dim f As String
    Let f = defaultDir & "\nailsPOS\setup\profile.txt"
    Open f For Random As #2 Len = Len(storeInfo)
    Get #2, 1, storeInfo
    Let storeName = Trim(storeInfo.name)
    Let storeAdd = Trim(storeInfo.street)
    Let storeCity = Trim(storeInfo.city)
    Let storeState = Trim(storeInfo.state)
    Let storeZip = storeInfo.zip
    Let storePhone = Trim(storeInfo.phone)
    Let storeWeb = Trim(storeInfo.web)
    Close #2
End Sub

Private Sub txtOldTranNum_GotFocus()
    Let txtOldTranNum.Text = ""
End Sub

Private Sub txtPassword_GotFocus()
    Let txtPassword.Text = ""
    Let txtPassword.PasswordChar = "*"
End Sub
