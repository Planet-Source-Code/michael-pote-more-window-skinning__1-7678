Attribute VB_Name = "Module2"
'You can just add this module to a project to skin it (Subject to rules described in form_load)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCPAINT = &HEE0086    ' dest = source OR dest
Private SS, Size As Boolean, Tx, Ty, Mov As Boolean, P As POINTAPI, Pressed() As Boolean

Sub SkinForm(Form1 As Form, WinPicHdc As Long, ButPicHdc As Long, BackPicHdc As Long, picBuff As PictureBox)
picBuff.Move 0, 0, Form1.ScaleWidth, 1 'Resize the buffer picture to the form width
StretchBlt picBuff.hdc, 0, 0, Form1.ScaleWidth, 1, BackPicHdc, 0, 0, 128, 1, SRCCOPY
'stretch the gradient on to the buffer image
picBuff.Picture = picBuff.Image
For i = 0 To Form1.ScaleHeight Step 1  'Screen Area
BitBlt Form1.hdc, 0, i, Form1.ScaleWidth, 1, picBuff.hdc, 0, 0, SRCCOPY
Next
For i = 0 To Form1.ScaleHeight Step 1 'Right Border
BitBlt Form1.hdc, Form1.ScaleWidth - 2, i, 2, 1, WinPicHdc, 126, 2, SRCCOPY
Next
For i = 0 To Form1.ScaleHeight Step 1 'Left Border
BitBlt Form1.hdc, 0, i, 2, 1, WinPicHdc, 0, 2, SRCCOPY
Next
For i = 0 To Form1.ScaleWidth Step 1 'Bottom Border
BitBlt Form1.hdc, i, Form1.ScaleHeight - 2, 1, 2, WinPicHdc, 2, 126, SRCCOPY
Next
For i = 0 To Form1.ScaleWidth - 3 Step 1 'Top Border
BitBlt Form1.hdc, i, 0, 1, 14, WinPicHdc, 3, 0, SRCCOPY
Next
BitBlt Form1.hdc, Form1.ScaleWidth - 2, 0, 2, 2, WinPicHdc, 126, 0, SRCCOPY
BitBlt Form1.hdc, 0, 0, 2, 14, WinPicHdc, 0, 0, SRCCOPY
BitBlt Form1.hdc, 0, Form1.ScaleHeight - 2, 2, 2, WinPicHdc, 0, 126, SRCCOPY
BitBlt Form1.hdc, Form1.ScaleWidth - 2, ScaleHeight - 2, 2, 2, WinPicHdc, 126, 126, SRCCOPY
BitBlt Form1.hdc, Form1.ScaleWidth - 3, 0, 2, 2, WinPicHdc, 126, 0, SRCCOPY
BitBlt Form1.hdc, Form1.ScaleWidth - 15, 2, 11, 11, WinPicHdc, 4, 32, SRCCOPY
Form1.FontBold = True
Form1.CurrentX = 5
Form1.CurrentY = 1
Form1.Print Form1.Caption
Form1.FontBold = False

For Each com In Form1.Controls 'Go thru all controls on form and check for buttons
If LCase(Mid(com.Name, 1, 4)) = "comm" Or LCase(Mid(com.Name, 1, 3)) = "cmd" Then
com.Height = 16
com.Visible = False
If com.Enabled = True Then
BitBlt Form1.hdc, com.Left, com.Top, 3, 16, ButPicHdc, 0, 0, SRCCOPY
BitBlt Form1.hdc, (com.Left + com.Width) - 3, com.Top, 3, 16, ButPicHdc, 45, 0, SRCCOPY
For i = com.Left + 3 To (com.Left + com.Width) - 3
BitBlt Form1.hdc, i, com.Top, 1, 16, ButPicHdc, 4, 0, SRCCOPY
Next
Else
BitBlt Form1.hdc, com.Left, com.Top, 3, 16, ButPicHdc, 0, 34, SRCCOPY
BitBlt Form1.hdc, (com.Left + com.Width) - 3, com.Top, 3, 16, ButPicHdc, 45, 34, SRCCOPY
For i = com.Left + 3 To (com.Left + com.Width) - 3
BitBlt Form1.hdc, i, com.Top, 1, 16, ButPicHdc, 4, 34, SRCCOPY
Next


End If
Form1.CurrentX = com.Left + 3
Form1.CurrentY = com.Top + 1
Form1.Print com.Caption
End If
Next
Form1.Refresh
End Sub


Sub MouseDown(Form1 As Form, ButPicHdc As Long, WinPicHdc As Long, X, Y)
With Form1
If X >= .ScaleWidth - 15 And Y <= 13 Then
BitBlt .hdc, .ScaleWidth - 15, 2, 11, 11, WinPicHdc, 4, 43, SRCCOPY
ElseIf Y <= 13 Then
Tx = X
Ty = Y
Mov = True
Else
If X >= .ScaleWidth - 5 Then Size = True: SS = "R"
If X <= 5 Then Size = True: SS = "L"
If Y >= .ScaleHeight - 5 Then Size = True: SS = "D"
If X >= .ScaleWidth - 5 And Y >= .ScaleHeight - 5 Then Size = True: SS = "DR"
For Each com In Form1.Controls
If LCase(Mid(com.Name, 1, 4)) = "comm" Or LCase(Mid(com.Name, 1, 3)) = "cmd" Then
If X >= com.Left And X <= com.Left + com.Width And Y >= com.Top And Y <= com.Top + com.Height And com.Enabled = True Then
Pressed(com.Index) = True
BitBlt .hdc, com.Left, com.Top, 3, 16, ButPicHdc, 0, 17, SRCCOPY
BitBlt .hdc, (com.Left + com.Width) - 3, com.Top, 3, 16, ButPicHdc, 45, 17, SRCCOPY
For i = com.Left + 3 To (com.Left + com.Width) - 3
BitBlt .hdc, i, com.Top, 1, 16, ButPicHdc, 4, 17, SRCCOPY
Next
.CurrentX = com.Left + 4
.CurrentY = com.Top + 2
Form1.Print com.Caption
End If
End If
Next
End If
.Refresh
End With
End Sub

Sub MouseMove(Form1 As Form, X, Y)
If X >= Form1.ScaleWidth - 5 Or X <= 5 Then
Form1.MousePointer = 9
If Y >= Form1.ScaleHeight - 5 Then Form1.MousePointer = 8
ElseIf Y >= Form1.ScaleHeight - 5 Then
Form1.MousePointer = 7
Else
Form1.MousePointer = 0
End If
If Mov Then
GetCursorPos P
DoEvents
Dim TempX As Long, TempY As Long
TempX = (P.X * 15) - (Tx * 15)
TempY = (P.Y * 15) - (Ty * 15)
Form1.Left = TempX
Form1.Top = TempY
End If
If Size Then
On Error Resume Next
GetCursorPos P
Select Case UCase(SS)
Case "R"
Form1.Width = (P.X * 15) - Form1.Left
Case "L"
Form1.Left = (P.X * 15)
Form1.Width = (P.X * 15)
Case "D"
Form1.Height = (P.Y * 15) - Form1.Top
Case "DR"
Form1.Width = (P.X * 15) - Form1.Left
Form1.Height = (P.Y * 15) - Form1.Top
End Select
End If
End Sub

Sub MouseUp(Form1 As Form, ButPicHdc As Long, X, Y)
Mov = False
Size = False
If X >= Form1.ScaleWidth - 15 And Y <= 13 Then
Form1.Hide
End If
For Each com In Form1.Controls
If LCase(Mid(com.Name, 1, 4)) = "comm" Or LCase(Mid(com.Name, 1, 3)) = "cmd" Then
If Pressed(com.Index) = True Then
BitBlt Form1.hdc, com.Left, com.Top, 3, 16, ButPicHdc, 0, 0, SRCCOPY
BitBlt Form1.hdc, (com.Left + com.Width) - 3, com.Top, 3, 16, ButPicHdc, 45, 0, SRCCOPY
For i = com.Left + 3 To (com.Left + com.Width) - 3
BitBlt Form1.hdc, i, com.Top, 1, 16, ButPicHdc, 4, 0, SRCCOPY
Next
Form1.CurrentX = com.Left + 2
Form1.CurrentY = com.Top + 1
Form1.Print com.Caption
ClickButton com.Index
Pressed(com.Index) = False
End If
End If
Next

End Sub

Sub FormLoad(Form1 As Form)
Form1.FontName = "Times New Roman"
Cnt = 0
For Each com In Form1.Controls
If LCase(Mid(com.Name, 1, 4)) = "comm" Or LCase(Mid(com.Name, 1, 3)) = "cmd" Then
Cnt = Cnt + 1
End If
Next
ReDim Pressed(0 To Cnt - 1) As Boolean
End Sub

Sub ClickButton(ind As Integer)
Select Case ind
Case 0
Form1.Skin = 0
Case 1
Form1.Skin = 1
Case 2
Form1.Command1(0).Enabled = (Form1.Command1(0).Enabled = False)
Form1.Command1(1).Enabled = (Form1.Command1(1).Enabled = False)
Form1.Command1(3).Enabled = (Form1.Command1(3).Enabled = False)
Form1.Command1(2).Caption = IIf(Form1.Command1(0).Enabled = True, "Disable", "Enable")
Case 3
Form1.Skin = 2
End Select
End Sub
