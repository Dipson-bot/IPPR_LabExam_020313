VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   16830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   8760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zooming by Intrepolation"
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zooming by replication"
      Height          =   735
      Left            =   5640
      TabIndex        =   6
      Top             =   6240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bit Plane Slicing"
      Height          =   855
      Left            =   5640
      TabIndex        =   5
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Intensity Level Slicing"
      Height          =   2535
      Left            =   1680
      TabIndex        =   2
      Top             =   5400
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "Without Background"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Background"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   5880
      ScaleHeight     =   3195
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   2160
      Picture         =   "Lab3 Intensity Level Slicing, Bit Plane Slicing and zooming.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
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

n = Text1.Text

If (n < 0) Then
n = 0
Text1.Text = 0
End If

If (n > 7) Then
n = 7
Text1.Text = 7
End If

If (r > (2 ^ n) - 1) Then
r = 255
Else
r = 0
End If
 
If (g > (2 ^ n) - 1) Then
g = 255
Else
g = 0
End If

If (b > (2 ^ n) - 1) Then
b = 255
Else
b = 0
End If

SetPixel Picture2.hdc, x, y, RGB(r, g, b)
Next y
Next x

End Sub

Private Sub Command2_Click()
Dim color As Long
Picture2.Height = Picture1.Height * 2
Picture2.Width = Picture1.Width * 2
For x = 0 To 180 * 2
For y = 0 To 180 * 2
color = GetPixel(Picture1.hdc, x / 2, y / 2)
SetPixel Picture2.hdc, x, y, color
Next y
Next x
End Sub

Private Sub Command4_Click()
Dim color As Long
Dim color1 As Long
Dim color2 As Long
Dim Gray As Long

For x = 0 To 100 * 2
For y = 0 To 100 * 2
If (x Mod 2 = 0) Then
color = GetPixel(Picture1.hdc, x / 2, y / 2)
Gray = color And 255
SetPixel Picture2.hdc, x, y, RGB(Gray, Gray, Gray)
End If
Next y
Next x

For x = 0 To 100 * 2
For y = 0 To 100 * 2
color1 = GetPixel(Picture1.hdc, x - 1, y) And 255
color2 = GetPixel(Picture2.hdc, x + 1, y) And 255
Gray = (color1 + color2) / 2
SetPixel Picture2.hdc, 2 * x, 2 * y + 1, RGB(Gray, Gray, Gray)
Next y
Next x

For x = 0 To 100 * 2
For y = 0 To 100 * 2
color1 = GetPixel(Picture1.hdc, x - 1, y) And 255
color2 = GetPixel(Picture2.hdc, x + 1, y) And 255
Gray = (color1 + color2) / 2
SetPixel Picture2.hdc, 2 * x + 1, 2 * y, RGB(Gray, Gray, Gray)
Next y
Next x
End Sub

Private Sub Option1_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
If r > 127 And r < 250 Then r = 256 - r Else r = r
If g > 127 And g < 250 Then g = 256 - g Else g = g
If b > 127 And b < 250 Then b = 256 - b Else b = b
SetPixel Picture2.hdc, x, y, RGB(r, g, b)
Next y
Next x
End Sub

Private Sub Option2_Click()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
If r > 127 And r < 250 Then r = 256 - r Else r = 0
If g > 127 And g < 250 Then g = 256 - g Else g = 0
If b > 127 And b < 250 Then b = 256 - b Else b = 0
SetPixel Picture2.hdc, x, y, RGB(r, g, b)
Next y
Next x
End Sub
