VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Clothes"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUnitPrice 
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdDeleteRecord 
      Caption         =   "Delete record"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuitProgram 
      Caption         =   "Quit the program"
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "Save record"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNewRecord 
      Caption         =   "Add new record"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtQuantity 
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtCategory 
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtClothNumber 
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtId 
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Discount"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "UnityPrice"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "ClothNumber"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As Object
Dim rs As Object

Private Sub Form_Load()
    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    ' Connection string for the Access database
    On Error GoTo ConnectionError
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\REVISION CORRECTION\BIT3\Market.mdb;Persist Security Info=False"
    MsgBox "WELCOME!", vbInformation, "WELCOME"
    Exit Sub

ConnectionError:
    MsgBox "Error connecting to the database: " & Err.Description, vbCritical, "Connection Error"
End Sub

Private Sub cmdAddNewRecord_Click()
    ' Clear all fields for a new record
    txtId.Text = ""
    txtName.Text = ""
    txtClothNumber.Text = ""
    txtCategory.Text = ""
    txtQuantity.Text = ""
    txtUnitPrice.Text = ""
    txtDiscount.Text = ""
    MsgBox "Fields cleared. Enter new record details.", vbInformation, "Add New Record"
End Sub

Private Sub cmdSaveRecord_Click()
    ' Save or update record
    On Error GoTo ErrorHandler

    Dim clothesID As String
    clothesID = txtId.Text

    ' Validate inputs before saving
    If txtName.Text = "" Or txtCategory.Text = "" Then
        MsgBox "Name and Category cannot be empty.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If Not IsNumeric(txtClothNumber.Text) Then
        MsgBox "Cloth Number must be a valid number.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If Not IsNumeric(txtQuantity.Text) Or Val(txtQuantity.Text) < 0 Then
        MsgBox "Quantity must be a valid non-negative number.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If Not IsNumeric(txtUnitPrice.Text) Or Val(txtUnitPrice.Text) < 0 Then
        MsgBox "Unit Price must be a valid non-negative number.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If Not IsNumeric(txtDiscount.Text) Or Val(txtDiscount.Text) < 0 Then
        MsgBox "Discount must be a valid non-negative number.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If clothesID = "" Then
        ' Insert new record
        conn.Execute "INSERT INTO Clothes (Name, ClothNumber, Category, Quantity, Unitprice, Discount) " & _
                     "VALUES ('" & txtName.Text & "', " & Val(txtClothNumber.Text) & ", '" & txtCategory.Text & "', " & _
                     Val(txtQuantity.Text) & ", " & Val(txtUnitPrice.Text) & ", " & Val(txtDiscount.Text) & ")"
        MsgBox "New record added successfully.", vbInformation, "Save Record"
    Else
        ' Update existing record
        conn.Execute "UPDATE Clothes SET Name = '" & txtName.Text & "', ClothNumber = " & Val(txtClothNumber.Text) & _
                     ", Category = '" & txtCategory.Text & "', Quantity = " & Val(txtQuantity.Text) & _
                     ", Unitprice = " & Val(txtUnitPrice.Text) & ", Discount = " & Val(txtDiscount.Text) & _
                     " WHERE Clothes_id = " & Val(clothesID)
        MsgBox "Record updated successfully.", vbInformation, "Save Record"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error saving record: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdDeleteRecord_Click()
    ' Delete record
    On Error GoTo ErrorHandler

    Dim clothesID As String
    clothesID = txtId.Text

    If clothesID = "" Then
        MsgBox "Please enter an ID to delete.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If Not IsNumeric(clothesID) Then
        MsgBox "Invalid ID format.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    conn.Execute "DELETE FROM Clothes WHERE Clothes_id = " & Val(clothesID)
    MsgBox "Record with ID " & clothesID & " deleted successfully.", vbInformation, "Delete Record"
    cmdAddNewRecord_Click ' Clear fields
    Exit Sub

ErrorHandler:
    MsgBox "Error deleting record: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdQuitProgram_Click()
    ' Close the application and ensure the connection is closed
    On Error Resume Next ' Handle any errors gracefully
    If conn.State = 1 Then
        conn.Close ' Only close if the connection is open
    End If
    Set conn = Nothing
    Set rs = Nothing
    Unload Me
End Sub

Private Sub cmdBack_Click()
    ' Navigate to the main form
    Form1.Show ' Replace "Form1" with the actual name of the main form
    Me.Hide
End Sub




Private Sub Text1_Change()

End Sub
