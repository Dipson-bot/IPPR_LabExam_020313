VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16065
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Gray Image"
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Vertical Flip"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Horizontal Flip"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Negative Image"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Image"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   3840
      Width           =   2655
   End
   Begin VB.PictureBox Picture6 
      Height          =   2895
      Left            =   9000
      ScaleHeight     =   2835
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   4920
      Width           =   3255
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   5520
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   4920
      Width           =   3135
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   1200
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   4680
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   9000
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   5040
      ScaleHeight     =   2715
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   1080
      Picture         =   "Lab1_Manipulation & processing of image.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   " Original Image"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture2.hdc, x, y, RGB(r, g, b)
Next y
Next x
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()
End Sub

Private Sub Command2_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture3.hdc, x, y, RGB(255 - r, 255 - g, 255 - b)
Next y
Next x
End Sub

Private Sub Command3_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
SetPixel Picture4.hdc, 255 - x, y, color
Next y
Next x
End Sub

Private Sub Command4_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
SetPixel Picture5.hdc, x, 255 - y, color
Next y
Next x
End Sub

Private Sub Command5_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
Grey = (r + g + b) / 3
SetPixel Picture6.hdc, x, y, RGB(Grey, Grey, Grey)
Next y
Next x
End Sub
