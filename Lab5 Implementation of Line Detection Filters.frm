VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Prewit Y Filter"
      Height          =   435
      Left            =   6600
      TabIndex        =   19
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Prewit X Filter"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sobel Y Filter"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sobel X Filter"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Laplacian Line Detection"
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-45 Line Detection"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+45 Line Detection"
      Height          =   615
      Left            =   9480
      TabIndex        =   13
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vertical Line Detection"
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Horizontal Line Detection"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.PictureBox Picture10 
      Height          =   2055
      Left            =   6600
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   7440
      Width           =   2055
   End
   Begin VB.PictureBox Picture9 
      Height          =   2055
      Left            =   3960
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.PictureBox Picture8 
      Height          =   2055
      Left            =   9480
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture7 
      Height          =   2055
      Left            =   6360
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture6 
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   720
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   9360
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   6720
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   3720
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   480
      Picture         =   "Lab5 Implementation of Line Detection Filters.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Original Image"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Horizontal Line Detection'
'-1 -1 -1'
' 2  2  2'
'-1 -1 -1'
Private Sub Command1_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = 2 * r
gg = 2 * g
bb = 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b


If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture2.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x

End Sub
'For Vertical Line Detection'
'-1  2 -1'
'-1  2 -1'
'-1  2 -1'
Private Sub Command2_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = 2 * r
gg = 2 * g
bb = 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture3.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x

End Sub

'For +45 Line Detection'
'-1  -1   2'
'-1   2  -1'
' 2  -1  -1'
Private Sub Command3_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = 2 * r
gg = 2 * g
bb = 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture4.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub
'For -45 Line Detection'
' 2  -1  -1'
'-1   2  -1'
'-1  -1   2'
Private Sub Command4_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = 2 * r
gg = 2 * g
bb = 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture5.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub

Private Sub Command5_Click()
'For Laplacian'
' 0  -1   0'
'-1   4  -1'
' 0  -1   0'
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = 4 * r
gg = 4 * g
bb = 4 * b


Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture6.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub
'For Sobel X-Filter'
'-1  0  1'
'-2  0  2'
'-1  0  1'
Private Sub Command6_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132


Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - 2 * r
gg = gg - 2 * g
bb = bb - 2 * b

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture7.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub

Private Sub Command7_Click()
'For Sobel Y-Filter'
'-1  -2  -1'
' 0   0   0'
' 1   2   1'
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - 2 * r
gg = gg - 2 * g
bb = bb - 2 * b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b


If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture8.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub

'For Prewit X Filter'
'-1  0  1'
'-1  0  1'
'-1  0  1'
Private Sub Command8_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132


Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x + 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b


Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x - 1, y)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b


If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture9.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x

End Sub

'For Prewit Y Filter'
'-1 -1 -1'
' 0  0  0'
' 1  1  1'
Private Sub Command9_Click()
Dim Color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 132
For y = 0 To 132

Color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr - r
gg = gg - g
bb = bb - b

Color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b

Color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = Color Mod 256
Color = Color / 256
g = Color Mod 256
Color = Color / 256
b = Color
rr = rr + r
gg = gg + g
bb = bb + b


If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture10.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x

End Sub

