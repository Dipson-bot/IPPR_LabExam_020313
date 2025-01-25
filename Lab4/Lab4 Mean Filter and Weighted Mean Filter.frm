VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Weighted Mean Filter"
      Height          =   855
      Left            =   6120
      TabIndex        =   3
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mean Filter"
      Height          =   975
      Left            =   2160
      TabIndex        =   2
      Top             =   5280
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   5640
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   1800
      Picture         =   "Lab4 Mean Filter and Weighted Mean Filter.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 180
For y = 0 To 180

color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x + 1, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x - 1, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

rr = rr * (1 / 9)
gg = gg * (1 / 9)
bb = bb * (1 / 9)

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture2.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x
End Sub

Private Sub Command2_Click()
Dim color As Long
Dim rr As Long
Dim gg As Long
Dim bb As Long
For x = 0 To 180
For y = 0 To 180

color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = 4 * r
gg = 4 * g
bb = 4 * b

color = GetPixel(Picture1.hdc, x - 1, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

color = GetPixel(Picture1.hdc, x + 1, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

color = GetPixel(Picture1.hdc, x, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

color = GetPixel(Picture1.hdc, x, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + 2 * r
gg = gg + 2 * g
bb = bb + 2 * b

color = GetPixel(Picture1.hdc, x - 1, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x + 1, y + 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x + 1, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b

color = GetPixel(Picture1.hdc, x - 1, y - 1)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
rr = rr + r
gg = gg + g
bb = bb + b


rr = rr * (1 / 16)
gg = gg * (1 / 16)
bb = bb * (1 / 16)

If (rr < 0) Then rr = 0 Else If rr > 255 Then rr = 255
If (gg < 0) Then gg = 0 Else If gg > 255 Then gg = 255
If (bb < 0) Then bb = 0 Else If bb > 255 Then bb = 255

SetPixel Picture2.hdc, x, y, RGB(rr, gg, bb)
Next y
Next x

End Sub
