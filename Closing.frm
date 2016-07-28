VERSION 5.00
Begin VB.Form Closing 
   Caption         =   "Closing Today's Transaction"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmClosingCash 
      Height          =   3495
      Left            =   600
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdCancelCash 
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
         TabIndex        =   14
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton cmdClearCash 
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
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdOkCash 
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
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtClosingCash 
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
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblClosingCash 
         Caption         =   "Enter Closing Cash"
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
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frmSure 
      Height          =   3375
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdSureNo 
         Caption         =   "No"
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
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdSureYes 
         Caption         =   "Yes"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblSure 
         Alignment       =   2  'Center
         Caption         =   "Are you sure?  You want to close today's Business."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame frmPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
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
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblPassword 
         Caption         =   "Enter Password"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Closing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Let txtPassword.Text = ""
    Closing.Hide
End Sub

Private Sub cmdCancelCash_Click()
    txtClosingCash.Text = ""
    frmClosingCash.Visible = False
    frmPassword.Visible = True
    Closing.Hide
End Sub

Private Sub cmdClearCash_Click()
    Let txtClosingCash.Text = ""
End Sub

Private Sub cmdOK_Click()
    Dim passwd As String
    
    Let passwd = Trim(txtPassword.Text)
    
    If UCase(passwd) <> "YES" Then
        Let txtPassword.PasswordChar = ""
        Let txtPassword.Text = "Wrong! Try again."
        Exit Sub
    End If
    txtPassword.Text = ""
    frmSure.Visible = True
    frmPassword.Visible = False
   
    
End Sub

Private Sub cmdOkCash_Click()
    Dim totalCommision As Single
    Dim record As Integer
    Dim item As salesItem
    Dim tSales As Single
    Dim tTax As Single
    Dim subTotal As Single
    Dim cashTotal As Single
    Dim ccTotal As Single
    Dim gcRedemTotal As Single
    Dim gcTotal As Single
    Dim changeTotal As Single
    Dim ccTips As Single
    Dim rCash As Single
    Dim ts As Single
    Dim sCash As Single
    Dim d As String
    Dim t As String
    
    Let d = Date
    Let t = Time
    
    If IsNumeric(Trim(txtClosingCash.Text)) = True Then
        Let rCash = Val(Trim(txtClosingCash.Text))
    Else
        Let txtClosingCash.Text = "Enter a number."
        Exit Sub
    End If
    
    Dim startclosingcash As openingCash
    
    Open defaultDir & "\nailsPOS\setup\openingCash.txt" For Random As #3 Len = Len(startclosingcash)
    Get #3, 1, startclosingcash
    Let sCash = startclosingcash.registerCash
    Let startclosingcash.registerCash = rCash
    Put #3, 1, startclosingcash
    Close #3
    
    Call PrintHead(d, t, 0)
    
    For record = 1 To recordNum
        Get #1, record, item
        
        If IsNumeric(Trim(item.itemType)) Then
            If Trim(item.itemType) = 101 Then
                Call PrintReceipt("Gift Certificate", item.price)
                Let gcTotal = gcTotal + item.price
            ElseIf Trim(item.itemType) = 102 Then
                Let subTotal = subTotal + item.price
            ElseIf Trim(item.itemType) = 103 Then
                Let tTax = tTax + item.price
            ElseIf Trim(item.itemType) = 104 Then
                Let ts = item.price
                Let tSales = tSales + item.price
            ElseIf Trim(item.itemType) = 105 Then
                Let ts = ts - item.price
                Let gcRedemTotal = gcRedemTotal + item.price
            ElseIf Trim(item.itemType) = 106 Then
                Let cashTotal = cashTotal + ts
            ElseIf Trim(item.itemType) = 107 Then
                Let ccTotal = ccTotal + ts
            ElseIf item.itemType = 108 Then
                Let changeTotal = changeTotal + item.price
            ElseIf item.itemType = 109 Then
                Let ccTips = ccTips + item.price
            End If
        Else
            Let totalCommision = totalCommision + item.commision
            Call PrintReceipt(Trim(item.name), item.price)
        End If
    Next record
    Printer.Print "Closing Record: Transaction List"
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Total sales:", subTotal)
    Call PrintReceipt("Total Sales Tax:", tTax)
    Call PrintReceipt("Gift Certificate Redemed:", gcRedemTotal * -1)
    Call PrintReceipt("Total Commision:", totalCommision * -1)
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Today's Net Amount:", tSales - gcRedemTotal - totalCommision)
    Call PrintReceipt("Starting Cash:", sCash)
    Printer.Print "--------------------------------------------------------"
    Call PrintReceipt("Total Amount in Register:", tSales - gcRedemTotal - totalCommision + sCash)
    Printer.Print
    Printer.Print
    Call PrintReceipt("Gift Certificate Sold:", gcTotal)
    Call PrintReceipt("Total Cash in Register:", cashTotal - totalCommision - ccTips)
    Call PrintReceipt("Total Credit Card Charge:", ccTotal + ccTips)
    Call PrintReceipt("Total Credit Card Tips:", ccTips)
  
    Printer.Print
    Printer.Print
    Printer.Print "                             ***End*** "
    Printer.EndDoc
    
    Close #1
    Name defaultDir & "\nailsPOS\sales\" & fileName As defaultDir & "\nailsPOS\sales\closed" & fileName
    
    
    frmClosingCash.Visible = False
    frmPassword.Visible = True
    Closing.Hide
    Register.Hide
    POS.Show
End Sub

Private Sub cmdSureNo_Click()
    frmPassword.Visible = True
    frmSure.Visible = False
    Closing.Hide
End Sub

Private Sub cmdSureYes_Click()
    frmSure.Visible = False
    frmClosingCash.Visible = True
    txtClosingCash.SetFocus
 
End Sub

Private Sub txtClosingCash_GotFocus()
    Let txtClosingCash.Text = ""
End Sub

Private Sub txtPassword_GotFocus()
    Let txtPassword.Text = ""
    Let txtPassword.PasswordChar = "*"
End Sub
