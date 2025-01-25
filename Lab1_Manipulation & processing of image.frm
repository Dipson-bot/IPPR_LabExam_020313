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
      Left            =   6960
      TabIndex        =   11
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Vertical Flip"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Horizontal Flip"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Negative Image"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Image"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   3720
      Width           =   2655
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   7200
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   4200
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   1200
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   7080
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   4080
      ScaleHeight     =   133
      ScaleMode       =   0  'User
      ScaleWidth      =   137.814
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   1080
      Picture         =   "Lab1_Manipulation & processing of image.frx":0000
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   " Original Image"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
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
For x = 0 To 132
For y = 0 To 132
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
For x = 0 To 132
For y = 0 To 132
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
For x = 0 To 132
For y = 0 To 132
color = GetPixel(Picture1.hdc, x, y)
SetPixel Picture4.hdc, 132 - x, y, color
Next y
Next x
End Sub

Private Sub Command4_Click()
Dim color As Long
For x = 0 To 132
For y = 0 To 132
color = GetPixel(Picture1.hdc, x, y)
SetPixel Picture5.hdc, x, 132 - y, color
Next y
Next x
End Sub

Private Sub Command5_Click()
Dim color As Long
For x = 0 To 132
For y = 0 To 132
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

