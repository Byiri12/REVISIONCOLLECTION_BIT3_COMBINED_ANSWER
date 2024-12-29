VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "                                                                                                             Bank Account Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7200
      ScaleHeight     =   1635
      ScaleWidth      =   3795
      TabIndex        =   28
      Top             =   5280
      Width           =   3855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   7200
      TabIndex        =   27
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7200
      TabIndex        =   26
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quit The Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   25
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove The Customer From The Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   24
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enter The Next Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   23
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add The Customer to The Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   22
      Top             =   8040
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   720
      Width           =   4095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Fixed Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Current Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   18
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bank Names"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   17
      Top             =   6480
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "BK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "BPR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   14
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Customers and Their Accounts"
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
      Left            =   7200
      TabIndex        =   20
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "Amount Deposited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Phone number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Lastname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Firstname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Collection to store customer information
Dim customers As Collection

Private Sub Form_Load()
    ' Initialize the customers collection
    Set customers = New Collection
End Sub

Private Sub Command1_Click()
    ' Add the customer to the ComboBox, ListBox, PictureBox, and TextBox
    Dim newCustomer As String
    Dim selectedBanks As String
    Dim accountType As String

    ' Validate inputs
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Then
        MsgBox "Please fill in all the required fields.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If Not IsNumeric(Text3.Text) Or Val(Text3.Text) <= 0 Then
        MsgBox "Please enter a valid age.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    If Not IsNumeric(Text7.Text) Or Val(Text7.Text) <= 0 Then
        MsgBox "Please enter a valid amount.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' Determine account type
    If Option1.Value = True Then
        accountType = "BPR"
    ElseIf Option2.Value = True Then
        accountType = "BK"
    Else
        MsgBox "Please select an account type.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' Validate and get selected banks
    selectedBanks = ""
    If Check1.Value = 1 Then ' Correctly check if Check1 is selected
        selectedBanks = selectedBanks & "Current Account, "
    End If
    If Check2.Value = 1 Then ' Correctly check if Check2 is selected
        selectedBanks = selectedBanks & "Fixed Account, "
    End If

    ' Check if no bank was selected
    If selectedBanks = "" Then
        MsgBox "Please select at least one bank.", vbExclamation, "Validation Error"
        Exit Sub
    End If

    ' Remove the trailing comma and space
    selectedBanks = Left(selectedBanks, Len(selectedBanks) - 2)

    ' Compile customer information
    newCustomer = "Name: " & Text1.Text & " " & Text2.Text & vbCrLf & _
                  "Age: " & Text3.Text & vbCrLf & _
                  "Gender: " & Text4.Text & vbCrLf & _
                  "ID: " & Text5.Text & vbCrLf & _
                  "Phone: " & Text6.Text & vbCrLf & _
                  "Amount: $" & Text7.Text & vbCrLf & _
                  "Account Type: " & accountType & vbCrLf & _
                  "Banks: " & selectedBanks

    ' Add customer to collection
    customers.Add newCustomer

    ' Display customer in ComboBox
    Combo1.AddItem Text1.Text & " " & Text2.Text

    ' Display customer in ListBox
    List1.AddItem newCustomer

    ' Display customer in PictureBox (use it as a label display area)
    Picture1.Cls ' Clear the PictureBox
    Picture1.FontSize = 10
    Picture1.Print newCustomer

    ' Display customer in a TextBox
    Text8.Text = newCustomer

    ' Clear input fields for the next customer
    ClearFields
End Sub



Private Sub Command2_Click()
    ' Clear all input fields for the next customer
    ClearFields
End Sub

Private Sub Command3_Click()
    ' Remove the selected customer from the ComboBox
    If Combo1.ListIndex >= 0 Then
        Dim selectedCustomer As String
        Dim i As Integer
        
        ' Get the selected customer from the ComboBox
        selectedCustomer = Combo1.Text

        ' Find and remove the customer from the collection
        For i = 1 To customers.Count
            If InStr(customers(i), selectedCustomer) > 0 Then
                customers.Remove i
                Exit For
            End If
        Next i

        ' Remove the customer from the ComboBox
        Combo1.RemoveItem Combo1.ListIndex

        ' Update the ListBox to reflect remaining customers
        List1.Clear
        For i = 1 To customers.Count
            List1.AddItem customers(i)
        Next i

        ' Optionally update the PictureBox and TextBox
        If customers.Count > 0 Then
            Picture1.Cls
            Picture1.FontSize = 10
            Picture1.Print customers(1) ' Show the first customer in the collection
            
            Text8.Text = customers(1) ' Display the first customer in the TextBox
        Else
            ' Clear displays if no customers remain
            Picture1.Cls
            Text8.Text = ""
        End If

        MsgBox "Customer removed successfully.", vbInformation, "Success"
    Else
        MsgBox "Please select a customer to remove.", vbExclamation, "Selection Error"
    End If
End Sub

Private Sub Command4_Click()
    ' Quit the application
    Unload Me
End Sub

Private Sub ClearFields()
    ' Clear all input fields
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Option1.Value = False
    Option2.Value = False
    Check1.Value = False
    Check2.Value = False
End Sub




