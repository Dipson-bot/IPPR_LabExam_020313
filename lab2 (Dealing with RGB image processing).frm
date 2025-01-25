VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   16080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   4440
      TabIndex        =   2
      Top             =   6600
      Width           =   4815
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Brightness"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Green"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Red"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   7200
      ScaleHeight     =   3195
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   2520
      Picture         =   "lab2 (Dealing with RGB image processing).frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture2.hdc, x, y, RGB(r + HScroll1.Value, g, b)
Next y
Next x
End Sub

Private Sub HScroll2_Change()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture2.hdc, x, y, RGB(r, g + HScroll2.Value, b)
Next y
Next x
End Sub

Private Sub HScroll3_Change()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture2.hdc, x, y, RGB(r, g, b + HScroll3.Value)
Next y
Next x
End Sub

Private Sub HScroll4_Change()
Dim color As Long
For x = 0 To 180
For y = 0 To 180
color = GetPixel(Picture1.hdc, x, y)
r = color Mod 256
color = color / 256
g = color Mod 256
color = color / 256
b = color
SetPixel Picture2.hdc, x, y, RGB(r + HScroll4.Value, g + HScroll4.Value, b + HScroll4.Value)
Next y
Next x
End Sub

