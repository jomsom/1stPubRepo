VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GiftCertificate 
   Caption         =   "Gift Certificate"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSellGiftcertificate 
      Caption         =   "Sell Gift Certificate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton cmdClose 
         Caption         =   "Done"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtId 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton cmdAddGC 
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtAmount 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid msgGiftcertificate 
         Height          =   3375
         Left            =   4800
         TabIndex        =   1
         Top             =   480
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   20
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Label lblNo 
         Caption         =   "Number"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame frmRedemGC 
      Caption         =   "Redem Gift Certificate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   8295
      Begin VB.PictureBox picMessage 
         Height          =   3615
         Left            =   4440
         ScaleHeight     =   3555
         ScaleWidth      =   3555
         TabIndex        =   16
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdRedemClose 
         Caption         =   "Done"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton cmdRedemClear 
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
         Left            =   1560
         TabIndex        =   14
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton cmdRedemOK 
         Caption         =   "OK"
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtRedemId 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtRedemAmount 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Number"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblRedemAmount 
         Caption         =   "Amount"
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
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "GiftCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gcNum As Integer
Dim tIndex As Integer

Private Sub cmdAddGC_Click()
     
    msgGiftcertificate.Clear
      
    Let msgGiftcertificate.row = 0
    Let msgGiftcertificate.Col = 0
    Let msgGiftcertificate.CellFontUnderline = True
    Let msgGiftcertificate.CellFontBold = True
    Let msgGiftcertificate.Text = Format("GiftCertificate No.", "!@@@@@@@@@@@@@@@@@@@@@")
    Let msgGiftcertificate.Col = 1
    Let msgGiftcertificate.CellFontUnderline = True
    Let msgGiftcertificate.CellFontBold = True
    Let msgGiftcertificate.Text = Format("Amount", "@@@@@@@@@")
    
    If IsNumeric(txtId.Text) And IsNumeric(txtAmount.Text) Then
        If Val(Trim(txtAmount.Text)) > 0 And (checkGCId() = False) Then
            Let tIndex = tIndex + 1
            Let gcAmount(tIndex) = Val(Trim(txtAmount.Text))
            Let gcId(tIndex) = Val(Trim(txtId.Text))
        ElseIf Val(Trim(txtAmount.Text)) <= 0 Then
            Let txtAmount.Text = "Enter again"
        ElseIf checkGCId() = True Then
        
        End If
    Else
        If IsNumeric(txtId.Text) = False Then
            Let txtId.Text = "Enter a number"
        Else
            Let txtAmount.Text = "Enter a number"
        End If
    End If
    
    For i = 1 To tIndex
        Let msgGiftcertificate.row = i
        Let msgGiftcertificate.Col = 0
        Let msgGiftcertificate.Text = gcId(i)
        Let msgGiftcertificate.Col = 1
        Let msgGiftcertificate.Text = gcAmount(i)
    Next i
    
End Sub

Private Sub cmdCancel_Click()
    msgGiftcertificate.Clear
    For i = 1 To tIndex
        Let gcAmount(i) = 0
        Let gcId(i) = 0
    Next i
    Let tIndex = 0
    Let txtId.Text = ""
    Let txtAmount.Text = ""
End Sub

Private Sub cmdClose_Click()
    Let tIndex = 0
    Let txtId.Text = ""
    Let txtAmount.Text = ""
    msgGiftcertificate.Clear
    GiftCertificate.Hide
   
End Sub

Private Sub cmdRedemOK_Click()
    If IsNumeric(txtRedemId.Text) And IsNumeric(txtRedemAmount.Text) Then
        If Val(Trim(txtRedemAmount.Text)) > 0 And (checkGCId() = False) Then
            Let tIndex = tIndex + 1
            Let gcAmount(tIndex) = Val(Trim(txtRedemAmount.Text))
            Let gcId(tIndex) = Val(Trim(txtRedemId.Text))
        ElseIf Val(Trim(txtRedemAmount.Text)) <= 0 Then
            Let txtRedemAmount.Text = "Enter again"
        ElseIf checkGCId() = True Then
        
        End If
    Else
        If IsNumeric(txtId.Text) = False Then
            Let txtId.Text = "Enter a number"
        Else
            Let txtAmount.Text = "Enter a number"
        End If
    End If
End Sub

Public Function checkGCId() As Boolean
    Dim gc As gcItem
    
    Open defaultDir & "\nailsPOS\gc\gc.txt" For Random As #2 Len = Len(gc)
    Let gcNum = LOF(2) / Len(gc)
    
    For i = 1 To gcNum
        Get #2, i, gc
        If Val(Trim(txtId.Text)) = gc.id Then
            checkGCId = True
            Let txtId.Text = "Already Sold"
            Close #2
            Exit Function
        End If
    Next i
    For i = 1 To tIndex
        If Val(Trim(txtId.Text)) = gcId(i) Then
            checkGCId = True
            Let txtId.Text = "Duplicate number"
            Close #2
            Exit Function
        End If
    Next i
    For i = 1 To 20
        If Val(Trim(txtId.Text)) = giftCertId(i) Then
            checkGCId = True
            Let txtId.Text = "Duplicate number"
            Close #2
            Exit Function
        End If
    Next i
    
    checkGCId = False
    Close #2
End Function

Private Sub msgGiftcertificate_DblClick()
    For i = msgGiftcertificate.row To tIndex
        Let gcAmount(i) = gcAmount(i + 1)
        Let gcId(i) = gcId(i + 1)
    Next i
    
    msgGiftcertificate.Clear
      
    Let msgGiftcertificate.row = 0
    Let msgGiftcertificate.Col = 0
    Let msgGiftcertificate.CellFontUnderline = True
    Let msgGiftcertificate.CellFontBold = True
    Let msgGiftcertificate.Text = Format("GiftCertificate No.", "!@@@@@@@@@@@@@@@@@@@@@")
    Let msgGiftcertificate.Col = 1
    Let msgGiftcertificate.CellFontUnderline = True
    Let msgGiftcertificate.CellFontBold = True
    Let msgGiftcertificate.Text = Format("Amount", "@@@@@@@@@")
    Let tIndex = tIndex - 1
    For i = 1 To tIndex
        Let msgGiftcertificate.row = i
        Let msgGiftcertificate.Col = 0
        Let msgGiftcertificate.Text = gcId(i)
        Let msgGiftcertificate.Col = 1
        Let msgGiftcertificate.Text = gcAmount(i)
    Next i
End Sub

Private Sub txtAmount_GotFocus()
    Let txtAmount.Text = ""
End Sub

Private Sub txtId_GotFocus()
    Let txtId.Text = ""
End Sub
