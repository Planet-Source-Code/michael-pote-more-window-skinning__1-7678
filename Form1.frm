VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "This is a Skinned Form."
   ClientHeight    =   5715
   ClientLeft      =   2595
   ClientTop       =   1530
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Skin 3"
      Height          =   240
      Index           =   3
      Left            =   60
      TabIndex        =   14
      Top             =   1380
      Width           =   1305
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   30
      Index           =   2
      Left            =   3270
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   13
      Top             =   1425
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   2
      Left            =   2925
      Picture         =   "Form1.frx":0542
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   1905
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picWin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   2
      Left            =   4380
      Picture         =   "Form1.frx":1984
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   11
      Top             =   1950
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disable"
      Height          =   375
      Index           =   2
      Left            =   1500
      TabIndex        =   10
      Top             =   990
      Width           =   1260
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   3075
      Picture         =   "Form1.frx":5DC6
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   2790
      Picture         =   "Form1.frx":6288
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   1770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picWin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   1
      Left            =   4185
      Picture         =   "Form1.frx":76CA
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   2085
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Skin 2"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   1110
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Skin 1"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Text            =   "This is a Skinned Form"
      Top             =   510
      Width           =   2655
   End
   Begin VB.PictureBox picBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   2685
      Picture         =   "Form1.frx":BB0C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   1650
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   45
      Picture         =   "Form1.frx":CF4E
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   6525
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   3300
      Picture         =   "Form1.frx":D410
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picWin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   0
      Left            =   3945
      Picture         =   "Form1.frx":D8D2
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   2235
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Skin As Integer
Private Sub Form_Load()
'Things to remember when skinning your own forms:
'------------------------------------------------

' * have all the nessisary pictures (described below)

' * When adding buttons create all buttons with
'   the name "Command1" then create a control array.

' * Set the forms border style to 0 - none.

' * To program in buttons click events goto
'   Module2's ClickButton sub.

Skin = 0
'Skins are "cut out" from these images:

'picWin() this is an array of images holding the windows borders and titlebar.
'picBut() these are the picures holding the buttons
'picBack() these images hold the background gradient
'picBuff this image holds the stretched picback for pasting onto the form
FormLoad Form1 'this sub inizilizes the skinning
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this sub checks if buttons are pressed and if the form should be resized.
MouseDown Form1, picBut(Skin).hdc, picWin(Skin).hdc, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this does the actual resizing
MouseMove Form1, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this clears the resize flags and presses the buttons
MouseUp Form1, picBut(Skin).hdc, X, Y
'for SkinForm, look at form_resize
SkinForm Form1, picWin(Skin).hdc, picBut(Skin).hdc, PicBack(Skin).hdc, picBuff
End Sub

Private Sub Form_Resize()
'This is the main sub that draws everything to the screen
SkinForm Form1, picWin(Skin).hdc, picBut(Skin).hdc, PicBack(Skin).hdc, picBuff
End Sub

Private Sub Text1_Change()
Form1.Caption = Text1.Text
SkinForm Form1, picWin(Skin).hdc, picBut(Skin).hdc, PicBack(Skin).hdc, picBuff
End Sub
