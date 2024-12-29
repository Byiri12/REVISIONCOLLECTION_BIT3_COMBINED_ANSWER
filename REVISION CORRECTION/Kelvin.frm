VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Select Conversion"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   855
      Left            =   5280
      TabIndex        =   7
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   855
      Left            =   840
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Conversion"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "to Rankin"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "to Kelvin"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Converted temperature"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Temperature"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Variable to store the input temperature and the result
    Dim temperature As Double
    Dim convertedTemperature As Double

    ' Get the temperature entered in Text1
    temperature = Val(Text1.Text)

    ' Check which option button is selected and perform the conversion
    If Option1.Value = True Then
        ' Convert to Kelvin
        convertedTemperature = ConvertToKelvin(temperature)
    ElseIf Option2.Value = True Then
        ' Convert to Rankin
        convertedTemperature = ConvertToRankin(temperature)
    Else
        MsgBox "Please select a conversion option.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Display the result in Text Box
    Text2.Text = convertedTemperature

  

    ' Display the result in Label
    Label2.Caption = "Converted Temperature: " & convertedTemperature
End Sub

' Function to convert to Kelvin
Private Function ConvertToKelvin(temp As Double) As Double
    ConvertToKelvin = temp
End Function

' Function to convert to Rankin
Private Function ConvertToRankin(temp As Double) As Double
    ConvertToRankin = temp * 1.8
End Function

Private Sub Command2_Click()
    ' Close the form
    Unload Me
End Sub
