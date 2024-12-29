VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   3720
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FULL NAME"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Function to get the first name
Private Function GetFirstname() As String
    GetFirstname = InputBox("Enter your first name:", "First Name")
End Function

' Function to get the last name
Private Function GetLastName() As String
    GetLastName = InputBox("Enter your last name:", "Last Name")
End Function

Private Sub Command1_Click()
    Dim firstName As String
    Dim lastName As String
    Dim fullName As String

    ' Get the first and last name using functions
    firstName = GetFirstname()
    lastName = GetLastName()

    ' Combine them to form the full name
    fullName = firstName & " " & lastName

    ' Display the full name in the PictureBox
    Picture1.Cls ' Clear the PictureBox before displaying
    Picture1.Print fullName
End Sub

