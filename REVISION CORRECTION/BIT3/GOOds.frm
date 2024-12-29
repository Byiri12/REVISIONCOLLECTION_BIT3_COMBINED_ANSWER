VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Goods"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   10800
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2760
      Top             =   8160
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"GOOds.frx":0000
      OLEDBString     =   $"GOOds.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtDiscount 
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToCloths 
      Caption         =   "Go to Clothes"
      Height          =   615
      Left            =   9120
      TabIndex        =   16
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeleteRecord 
      Caption         =   "Delete the program"
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitProgram 
      Caption         =   "Quit the program"
      Height          =   615
      Left            =   5520
      TabIndex        =   14
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "Save record"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddNewRecord 
      Caption         =   "Add new record"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox txtQuantity 
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtCategory 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtItemNumber 
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtId 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Discount"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "UnityPrice"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "ItemNumber"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
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
    MsgBox "WELCOME!.", vbInformation, "WELCOME"
    Exit Sub

ConnectionError:
    MsgBox "Error connecting to the database: " & Err.Description, vbCritical, "Connection Error"
End Sub

Private Sub cmdAddNewRecord_Click()
    ' Clear all fields for a new record
    txtId.Text = ""
    txtName.Text = ""
    txtItemNumber.Text = ""
    txtCategory.Text = ""
    txtQuantity.Text = ""
    txtUnitPrice.Text = ""
    txtDiscount.Text = ""
    MsgBox "Fields cleared. Enter new record details.", vbInformation, "Add New Record"
End Sub

Private Sub cmdSaveRecord_Click()
    ' Save or update record
    On Error GoTo ErrorHandler

    Dim goodsID As String
    goodsID = txtId.Text

    ' Validate inputs before saving
    If txtName.Text = "" Or txtCategory.Text = "" Then
        MsgBox "Name and Category cannot be empty.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    If Not IsNumeric(txtItemNumber.Text) Then
        MsgBox "Item Number must be a valid number.", vbExclamation, "Input Error"
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

    If goodsID = "" Then
        ' Insert new record
        Dim cmd As Object
        Set cmd = CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = "INSERT INTO Goods (Name, ItemNumber, Category, Quantity, Unitprice, Discount) VALUES (?, ?, ?, ?, ?, ?)"
        
        ' Add parameters
        cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, txtName.Text)  ' Name (Text)
        cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , Val(txtItemNumber.Text))  ' ItemNumber (Integer)
        cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, txtCategory.Text)  ' Category (Text)
        cmd.Parameters.Append cmd.CreateParameter(, 3, 1, , Val(txtQuantity.Text))  ' Quantity (Integer)
        cmd.Parameters.Append cmd.CreateParameter(, 5, 1, , Val(txtUnitPrice.Text))  ' UnitPrice (Double)
        cmd.Parameters.Append cmd.CreateParameter(, 5, 1, , Val(txtDiscount.Text))  ' Discount (Double)
        
        cmd.Execute
        MsgBox "New record added successfully.", vbInformation, "Save Record"
    Else
        ' Validate ID before updating
        If Not IsNumeric(goodsID) Then
            MsgBox "Invalid ID format.", vbExclamation, "Input Error"
            Exit Sub
        End If

        ' Update existing record
        Dim updateCmd As Object
        Set updateCmd = CreateObject("ADODB.Command")
        updateCmd.ActiveConnection = conn
        updateCmd.CommandText = "UPDATE Goods SET Name = ?, ItemNumber = ?, Category = ?, Quantity = ?, Unitprice = ?, Discount = ? WHERE goods_id = ?"

        ' Add parameters for update
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 8, 1, 255, txtName.Text)
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 3, 1, , Val(txtItemNumber.Text))
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 8, 1, 255, txtCategory.Text)
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 3, 1, , Val(txtQuantity.Text))
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 5, 1, , Val(txtUnitPrice.Text))
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 5, 1, , Val(txtDiscount.Text))
        updateCmd.Parameters.Append updateCmd.CreateParameter(, 3, 1, , Val(goodsID)) ' goods_id for the update
        
        updateCmd.Execute
        MsgBox "Record updated successfully.", vbInformation, "Save Record"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error saving record: " & Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "Line: " & Erl, vbCritical, "Error"
End Sub

Private Sub cmdDeleteRecord_Click()
    ' Delete record
    On Error GoTo ErrorHandler

    Dim goodsID As String
    goodsID = txtId.Text

    If goodsID = "" Then
        MsgBox "Please enter an ID to delete.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If Not IsNumeric(goodsID) Then
        MsgBox "Invalid ID format.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    conn.Execute "DELETE FROM Goods WHERE goods_id = " & Val(goodsID)
    MsgBox "Record with ID " & goodsID & " deleted successfully.", vbInformation, "Delete Record"
    cmdAddNewRecord_Click ' Clear fields
    Exit Sub

ErrorHandler:
    MsgBox "Error deleting record: " & Err.Description & vbCritical, "Error"
End Sub

Private Sub cmdQuitProgram_Click()
    ' Close the application and ensure the database connection is properly closed
    On Error Resume Next ' Handle any errors gracefully

    ' Check if the database connection is open, and close it
    If Not conn Is Nothing Then
        If conn.State = 1 Then ' Connection is open
            conn.Close ' Close the connection
        End If
        Set conn = Nothing ' Release the connection object
    End If

    ' Release the recordset object if used
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If

    ' Unload the current form
    Unload Me

    ' Optional: Terminate the application completely
    End
End Sub


Private Sub cmdGoToCloths_Click()
    ' Navigate to Form2 (Cloths)
    Form1.Show
    Me.Hide
End Sub

Private Sub cmdBack_Click()
    ' Navigate back to Form1 (Main form)
    Form1.Show
    Me.Hide
End Sub


