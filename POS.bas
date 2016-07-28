Attribute VB_Name = "SaleItem"
Public pCheck As Boolean
Public defaultDir As String
Public giftCertId(1 To 20) As Long
Public gcId(1 To 20) As Long
Public gcAmount(1 To 20) As Single

Public manicureName(1 To 12) As String
Public manicurePrice(1 To 12) As Single
Public pedicureName(1 To 10) As String
Public pedicurePrice(1 To 10) As Single
Public threadingName(1 To 8) As String
Public threadingPrice(1 To 8) As Single
Public waxingName(1 To 14) As String
Public waxingPrice(1 To 14) As Single
Public tipsName(1 To 10) As String
Public tipsPrice(1 To 10) As Single
Public massageName(1 To 6) As String
Public massagePrice(1 To 6) As Single
Public specialName(1 To 6) As String
Public specialPrice(1 To 6) As Single
Public othersName(1 To 8) As String
Public othersPrice(1 To 8) As Single
Public productName(1 To 40) As String
Public productPrice(1 To 40) As Single
Public discountName(1 To 4) As String
Public discountPrice(1 To 4) As Single

Public massageCommision(1 To 6) As Single
Public pedicureCommision(1 To 10) As Single
Public othersCommision(1 To 8) As Single
Public specialCommision(1 To 6) As Single

Rem *****************************
Public fileName As String
Public oldFileName As String
Public lastRow As Integer
Public dstr As String
Public tstr As String
Public itemArray(1 To 100) As Integer

Public tType As Integer
Public tenderAmt As Single
Public balance As Single

Public redemGCAmount As Single
Public redemGCArray(1 To 20) As Single
Public redemGCNum As Integer
Public giftCertSum As Single
Public giftCertNum As Integer
Public giftCertArray(1 To 20) As Single

Public itemIndex As Integer
Public recordNum As Integer
Public transactionNum As Integer
Public itemType(1 To 100) As String
Public itemCommision(1 To 100) As Single
Public itemPrice(1 To 100) As Single
Public itemName(1 To 100) As String
Public totalSales As Single, tax As Single, total As Single, subTotal As Single
Rem ******************************************************


Public voided As Integer
Public empNum As Integer
Public storeName As String
Public storeAdd As String
Public storeCity As String
Public storeState As String
Public storeZip As Integer
Public storePhone As String
Public storeWeb As String


Public Type employee
    name As String * 20
    id As Integer
End Type

Public Type openingCash
    registerCash As Single
End Type

Public Type gcItem
    amount As Single
    id As Long
    status As Integer
    dstr As String * 15
    tstr As String * 15
End Type

Public Type product
    name As String * 28
    price As Single
    productId As String * 5
    commision As Single
End Type


Public Type profile
    name As String * 30
    street As String * 30
    city As String * 20
    state As String * 20
    zip As Integer
    phone As String * 15
    web As String * 30
End Type

Public Type salesItem
    name As String * 28
    price As Single
    tranNum As Integer
    itemType As String * 5
    dateStr As String * 15
    timeStr As String * 15
    commision As Single
    
End Type

Public Function checkEmpName(name As String) As Boolean
    Dim emp As employee
    For i = 1 To empNum
        Get #2, i, emp
        If RTrim(emp.name) = UCase(name) Then
            checkEmpName = False
            Exit Function
        End If
    Next i
    checkEmpName = True
End Function

Public Sub PrintHead(d As String, t As String, tNum As Integer)
    
    Rem Printer.Print Space(83) & "h"
    
    Printer.Print Tab(18); storeName
    Printer.Print Tab(18); storeAdd
    Printer.Print Tab(18); storeCity & ", " & storeState & " " & storeZip
    Printer.Print Tab(18); storePhone
    Printer.Print Tab(18); storeWeb
    Printer.Print
    Printer.Print "Date: " & d
    Printer.Print
    Printer.Print "Time: " & t
    Printer.Print
    If (tNum > 0) Then
        Printer.Print "Transaction: " & tNum
        Printer.Print
        Printer.Print
        Printer.Print "ITEM"; Tab(37); "PRICE"
    End If
    Printer.Print "--------------------------------------------------------"
End Sub

Public Sub ClearMsg(msg As MSFlexGrid)
    For i = 1 To msg.Rows - 1
        Let msg.row = i
        For j = 0 To msg.Cols - 1
            Let msg.Col = j
            If (j = 1) Then
                If Trim(msg.Text) = "" Then
                    Exit Sub
                End If
            End If
            Let msg.Text = ""
        Next j
    Next i
End Sub
Public Sub PrintReceipt(str As String, num As Single)
    Dim fmt1 As String, fmt2 As String, fmt3 As String
    Dim col1 As String, col2 As String
    
    Let fmt1 = "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
    Let fmt2 = "0.00"
    Let fmt3 = "@@@@@@@"
    Let col1 = Format(str, fmt1)
    Let col2 = Format(num, fmt2)
    Let col2 = Format(col2, fmt3)

    If num < 10 And num > -0.01 Then
        Printer.Print col1; Tab(35); "   " & col2
    ElseIf num < 100 And num > 9.99 Then
        Printer.Print col1; Tab(35); "  " & col2
    ElseIf num > -10 And num < 0 Then
        Printer.Print col1; Tab(35); "   " & col2
    ElseIf num < 1000 And num > 99.99 Then
        Printer.Print col1; Tab(35); " " & col2
    ElseIf num > -100 And num < -9.99 Then
        Printer.Print col1; Tab(35); "  " & col2
    ElseIf num > 999.99 Then
        Printer.Print col1; Tab(35); col2
    Else
        Printer.Print col1; Tab(35); " " & col2
    End If
    
End Sub

Public Sub ResetSales()
    Dim i As Integer
    
    For i = 1 To 100
        Let itemPrice(i) = 0
        Let itemName(i) = ""
        Let itemType(i) = ""
        Let itemCommision(i) = 0
    Next i
    
    For i = 1 To 20
        Let giftCertArray(i) = 0
        Let giftCertId(i) = 0
    Next i
    
    For i = 1 To 20
        Let redemGCArray(i) = 0
    Next i
    
    Let tenderAmt = 0
    Let balance = 0
    Let redemGCNum = 0
    Let redemGCAmount = 0
    Let giftCertNum = 0
    Let giftCertSum = 0
    Let itemIndex = 0
    Let totalSales = 0
    Let subTotal = 0
    Let tax = 0
    Let total = 0
    Let lastRow = 0
    Let tType = 0
    
End Sub

Public Function formatString(str As String) As String
    Dim strLength As Integer
    Dim tempStr As String
    Let strLength = Register.TextWidth(str)
    For i = 1 To strLength
        Let tempStr = tempStr & "@"
    Next i
    formatString = tempStr
End Function

Public Function ChkPassword() As Boolean
    password.Show 1
    ChkPassword = pCheck
End Function
