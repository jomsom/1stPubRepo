VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ServiceTicket 
   Caption         =   "Service Ticket"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTicket 
      Caption         =   "Ticket Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdClose 
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
         Left            =   1080
         TabIndex        =   7
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
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
         Left            =   1080
         TabIndex        =   8
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Ticket"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtId 
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
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtName 
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
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid msgEmpList 
         Height          =   4935
         Left            =   5160
         TabIndex        =   1
         Top             =   480
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   50
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Label lblId 
         Caption         =   "ID"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblName 
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "ServiceTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    Let txtName.Text = ""
    Let txtId.Text = ""
End Sub

Private Sub cmdClose_Click()
    Let txtId.Text = ""
    Let txtName.Text = ""
    ServiceTicket.Hide
End Sub

Private Sub cmdPrint_Click()
    Dim dstr As String
    Dim tstr As String
    Dim p As profile
    Let dstr = Date
    Let tstr = Time
    
    
    Printer.Print "Time: " & tstr & "        Date: " & dstr
    
    Printer.Print "Name: " & txtName.Text & "     ID #" & txtId.Text
    Printer.Print
    Printer.Print "1. MANICURE: (Regular) (French) (SPA)  "
    Printer.Print
    Printer.Print
    Printer.Print "2. PEDICURE: (Regular) (French) (SPA)"
    Printer.Print
    Printer.Print
    Printer.Print "3. WAXING: (Eyebrow) (Lip) (Bikini) "
    Printer.Print "  (Brazillian: Mini Full)"
    Printer.Print
    Printer.Print
    Printer.Print "4. MASSAGE:  "
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print storeName
    Printer.Print "."
    Printer.EndDoc
    
    Let txtName.Text = ""
    Let txtId.Text = ""
    
End Sub

Private Sub Form_Load()
    Dim emp As employee
    
    Open defaultDir & "\nailsPOS\setup\employee.txt" For Random As #2 Len = Len(emp)
    Let empNum = LOF(2) / Len(emp)
    msgEmpList.Clear
    
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

Private Sub msgEmpList_DblClick()
    Let msgEmpList.Col = 0
    Let txtId.Text = msgEmpList.Text
    Let msgEmpList.Col = 1
    Let txtName.Text = msgEmpList.Text
End Sub



Private Sub txtId_GotFocus()
    Let txtId.Text = ""
End Sub

Private Sub txtName_GotFocus()
    Let txtName.Text = ""
End Sub
