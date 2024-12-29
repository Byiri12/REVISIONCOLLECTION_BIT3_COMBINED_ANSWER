VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Area of Triangle"
      Height          =   855
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   7200
      ScaleHeight     =   2355
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Calculate Area of Triangle"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Height"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Base"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Variables to store base, height, and area
    Dim baseValue As Double
    Dim heightValue As Double
    Dim triangleArea As Double

    ' Get the base and height values from Text Boxes
    baseValue = Val(Text1.Text)
    heightValue = Val(Text2.Text)
    
    ' Calculate the area using the Area function
    triangleArea = Area(baseValue, heightValue)
    
    ' Display the result in Text Box
    Text3.Text = triangleArea
    
    ' Display the result in Picture Box
    Picture1.Cls ' Clear Picture Box
    Picture1.Print "Area of Triangle: " & triangleArea
    
    ' Display the result in Label
    Label3.Caption = "Area of Triangle: " & triangleArea
End Sub

' User-defined function to calculate the area of a triangle
Private Function Area(base As Double, height As Double) As Double
    Area = 0.5 * base * height
End Function
