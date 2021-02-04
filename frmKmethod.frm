VERSION 5.00
Begin VB.Form frmK 
   AutoRedraw      =   -1  'True
   Caption         =   "Karnaugh Simplification Method"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "3-Input functions"
      Height          =   615
      Left            =   4080
      TabIndex        =   27
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   44
      Left            =   2880
      TabIndex        =   16
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   43
      Left            =   2160
      TabIndex        =   15
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   42
      Left            =   1440
      TabIndex        =   14
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   41
      Left            =   720
      TabIndex        =   13
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   34
      Left            =   2880
      TabIndex        =   12
      Text            =   "1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   33
      Left            =   2160
      TabIndex        =   11
      Text            =   "1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   32
      Left            =   1440
      TabIndex        =   10
      Text            =   "1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   31
      Left            =   720
      TabIndex        =   9
      Text            =   "1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   24
      Left            =   2880
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   23
      Left            =   2160
      TabIndex        =   7
      Text            =   "1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   22
      Left            =   1440
      TabIndex        =   6
      Text            =   "1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   21
      Left            =   720
      TabIndex        =   5
      Text            =   "1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   14
      Left            =   2880
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   13
      Left            =   2160
      TabIndex        =   3
      Text            =   "1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   12
      Left            =   1440
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtKV 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   11
      Left            =   720
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdSimplify 
      Caption         =   "Simplify"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Shape box2R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box1R1C 
      BackColor       =   &H8000000D&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box2R1C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box1R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box4R1C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Height          =   2895
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape box1R4C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      Height          =   735
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape box4R2C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2895
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape box2R4C 
      BackColor       =   &H8000000D&
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1455
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "10"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   25
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   195
      Index           =   4
      Left            =   2880
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   21
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   20
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      Caption         =   "DC"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      Caption         =   "BA"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   0
      Width           =   210
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   120
      Y1              =   600
      Y2              =   120
   End
End
Attribute VB_Name = "frmK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


' ***********************************************
' ** This program solves a 4x4 KV-diagram      **
' ** It even looks for random values indicated **
' ** by 'X' or another non-numeric symbol.     **
' ** Input: values in the KV-diagram           **
' ** Ouput: the simpelest equation possible.   **
' ***********************************************

Const vbWhite = &HFFFFFF
Const vbLightYellow = &H80FFFF
Const vbYellow = &HFFFF&
Const vbLightOrange = &H80C0FF
Const vbOrange = &H80FF&
Const vbLightGreen = &H80FF80
' vbGreen is a standard color
Const vbLightCyan = &HFFFF80
Const vbCyan = &HFFFF00
Const vbLightBlue = &HFF8080
' vbBlue is a standard color
Const vbLightPurple = &HFF80FF
Const vbPurple = &HFF00FF
Const vbMagenta = &HFF00FF
Dim Tested(1 To 4, 1 To 4) As String * 1

Private Sub cmdSimplify_Click()



 MsgBox F
End Sub

Private Sub SetValue(C1 As Byte, C2 As Byte)
 Tested(C1, C2) = "1"
End Sub

Public Function F() As String
 
 Dim C1 As Byte, C2 As Byte
 Dim C1a As Byte, C2a As Byte
 Dim Test As String, strTested As String, Bool As Boolean
 Dim Same As Boolean
 Dim Obj As Object
 ' Set tested to false
 For C1 = 1 To 4
  For C2 = 1 To 4
   Tested(C1, C2) = "0"
   txtKV(10 * C1 + C2).BackColor = vbWhite
   If txtKV(10 * C1 + C2) = "0" Then Tested(C1, C2) = "1"
  Next C2
 Next C1
 ' Unload all the gridlines
 For C1 = 1 To 4
  Select Case C1
   Case 1: Set Obj = box2R4C ' Mask8_Rows
   Case 2: Set Obj = box4R2C ' Mask8_Columns
   Case 3: Set Obj = box1R4C ' Mask4_Rows
   Case 4: Set Obj = box4R1C ' Mask4_Columns
   Case 5: Set Obj = box2R2C ' Mask4_Cubes
   Case 6: Set Obj = box1R2C ' Mask2_Rows
   Case 7: Set Obj = box2R1C ' Mask2_Columns
   Case 8: Set Obj = box1R1C ' Mask1
  End Select
  For C2 = 1 To Obj.Count - 1
   Unload Obj(C2)
  Next C2
 Next C1
Mask16:
 Test = txtKV(11) + txtKV(12) + txtKV(13) + txtKV(14)
 Test = Test + txtKV(21) + txtKV(22) + txtKV(23) + txtKV(24)
 Test = Test + txtKV(31) + txtKV(32) + txtKV(33) + txtKV(34)
 Test = Test + txtKV(41) + txtKV(42) + txtKV(43) + txtKV(44)
 If (InStr(Test, "0")) = 0 Then
  F = "1"
  GoTo EndFunction
 End If
Mask8_Rows:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  Test = txtKV(C1 * 10 + 1) + txtKV(10 * C1 + 2) + txtKV(10 * C1 + 3) + txtKV(10 * C1 + 4)
  strTested = Tested(C1, 1) + Tested(C1, 2) + Tested(C1, 3) + Tested(C1, 4)
  Test = Test + txtKV(C1a * 10 + 1) + txtKV(10 * C1a + 2) + txtKV(10 * C1a + 3) + txtKV(10 * C1a + 4)
  strTested = strTested + Tested(C1a, 1) + Tested(C1a, 2) + Tested(C1a, 3) + Tested(C1a, 4)
  Bool = (InStr(strTested, "0") <> 0) And (InStr(Test, "0") = 0)
  If Bool Then
   Call SetValue(C1, 1): Call SetValue(C1a, 1)
   Call SetValue(C1, 2): Call SetValue(C1a, 2)
   Call SetValue(C1, 3): Call SetValue(C1a, 3)
   Call SetValue(C1, 4): Call SetValue(C1a, 4)
   Set Obj = box2R4C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 * C1 + 1).Top - 120
    .Left = txtKV(10 * C1 + 1).Left - 120
    .Visible = True
   End With
   Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
   End If
   Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
   End If
  End If
 Next C1
Mask8_Columns:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  Test = txtKV(10 + C1) + txtKV(20 + C1) + txtKV(30 + C1) + txtKV(40 + C1)
  strTested = Tested(1, C1) + Tested(2, C1) + Tested(3, C1) + Tested(4, C1)
  Test = Test + txtKV(10 + C1a) + txtKV(20 + C1a) + txtKV(30 + C1a) + txtKV(40 + C1a)
  strTested = strTested + Tested(1, C1a) + Tested(2, C1a) + Tested(3, C1a) + Tested(4, C1a)
  Bool = InStr(strTested, "0") <> 0 And (InStr(Test, "0")) = 0
  If Bool Then
   Call SetValue(1, C1): Call SetValue(1, C1a)
   Call SetValue(2, C1): Call SetValue(2, C1a)
   Call SetValue(3, C1): Call SetValue(3, C1a)
   Call SetValue(4, C1): Call SetValue(4, C1a)
   Set Obj = box4R2C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 + C1).Top - 120
    .Left = txtKV(10 + C1).Left - 120
    .Visible = True
   End With
   Same = Not ((CBool(Val(Left$(lblC(C1), 1)))) Xor (CBool(Val(Left$(lblC(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C1), 1) = "0" Then F = F + "' "
   End If
   Same = Not ((CBool(Val(Right$(lblC(C1), 1)))) Xor (CBool(Val(Right$(lblC(C1a), 1)))))
   If Same Then
    If F <> Empty Then F = F + " + "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C1), 1) = "0" Then F = F + "' "
   End If
  End If
 Next C1
Mask4_Rows:
 For C1 = 1 To 4
  Test = txtKV(10 * C1 + 1) + txtKV(10 * C1 + 2) + txtKV(10 * C1 + 3) + txtKV(10 * C1 + 4)
  strTested = Tested(C1, 1) + Tested(C1, 2) + Tested(C1, 3) + Tested(C1, 4)
  Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
  If Bool Then
   Call SetValue(C1, 1): Call SetValue(C1, 2)
   Call SetValue(C1, 3): Call SetValue(C1, 4)
   Set Obj = box1R4C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 * C1 + 1).Top - 120
    .Left = txtKV(10 * C1 + 1).Left - 120
    .Visible = True
   End With
   If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
   If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
   If Right$(lblR(C1), 1) = "0" Then F = F + "' "
  End If
 Next C1
Mask4_columns:
 For C1 = 1 To 4
  Test = txtKV(10 + C1) + txtKV(20 + C1) + txtKV(30 + C1) + txtKV(40 + C1)
  strTested = Tested(1, C1) + Tested(2, C1) + Tested(3, C1) + Tested(4, C1)
  Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
  If Bool Then
   Call SetValue(1, C1): Call SetValue(2, C1)
   Call SetValue(3, C1): Call SetValue(4, C1)
   Set Obj = box4R1C
   Load Obj(Obj.Count)
   With Obj(Obj.Count - 1)
    .Top = txtKV(10 + C1).Top - 120
    .Left = txtKV(10 + C1).Left - 120
    .Visible = True
   End With
   If F <> Empty Then F = F + " + "
    F = F + Left$(lblC(0), 1)
   If Left$(lblC(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
   If Right$(lblC(C1), 1) = "0" Then F = F + "' "
  End If
 Next C1
Mask4_Cubes:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  For C2 = 1 To 4
   If C2 < 4 Then
    C2a = C2 + 1
   Else
    C2a = 1
   End If
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1 + C2a)
   strTested = Tested(C1, C2) + Tested(C1, C2a)
   Test = Test + txtKV(10 * C1a + C2) + txtKV(10 * C1a + C2a)
   strTested = strTested + Tested(C1a, C2) + Tested(C1a, C2a)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   If Bool Then
    Call SetValue(C1, C2): Call SetValue(C1a, C2)
    Call SetValue(C1, C2a): Call SetValue(C1a, C2a)
    Set Obj = box2R2C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = True
    End With
    If F <> Empty Then F = F + " + "
    Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
    If Same Then
     F = F + Left$(lblR(0), 1)
     If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
    If Same Then
     F = F + Right$(lblR(0), 1)
     If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Left$(lblC(C2), 1)))) Xor (CBool(Val(Left$(lblC(C2a), 1)))))
    If Same Then
     F = F + Left$(lblC(0), 1)
     If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblC(C2), 1)))) Xor (CBool(Val(Right$(lblC(C2a), 1)))))
    If Same Then
     F = F + Right$(lblC(0), 1)
     If Right$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
   End If
  Next C2
 Next C1
Mask2_Rows:
 For C1 = 1 To 4
  For C2 = 1 To 4
   If C2 < 4 Then
    C2a = C2 + 1
   Else
    C2a = 1
   End If
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1 + C2a)
   strTested = Tested(C1, C2) + Tested(C1, C2a)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   Bool = Bool And Not ((InStr("01", Left$(Test, 1)) = 0 And Tested(C1, C2a) = "1") Or (InStr("01", Right$(Test, 1)) = 0 And Tested(C1, C2) = "1"))
   If Bool Then
    Call SetValue(C1, C2): Call SetValue(C1, C2a)
    Set Obj = box1R2C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = True
    End With
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    
    Same = Not ((CBool(Val(Left$(lblC(C2), 1)))) Xor (CBool(Val(Left$(lblC(C2a), 1)))))
    If Same Then
     F = F + Left$(lblC(0), 1)
     If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblC(C2), 1)))) Xor (CBool(Val(Right$(lblC(C2a), 1)))))
    If Same Then
     F = F + Right$(lblC(0), 1)
     If Right$(lblC(C2), 1) = "0" Then F = F + "' "
    End If
   End If
  Next C2
 Next C1
Mask2_Columns:
 For C1 = 1 To 4
  If C1 < 4 Then
   C1a = C1 + 1
  Else
   C1a = 1
  End If
  For C2 = 1 To 4
   Test = txtKV(10 * C1 + C2) + txtKV(10 * C1a + C2)
   strTested = Tested(C1, C2) + Tested(C1a, C2)
   Bool = InStr(strTested, "0") <> 0 And InStr(Test, "0") = 0
   Bool = Bool And Not ((InStr("01", Left$(Test, 1)) = 0 And Tested(C1a, C2) = "1") Or (InStr("01", Right$(Test, 1)) = 0 And Tested(C1, C2) = "1"))
   If Bool Then
    Call SetValue(C1, C2)
    Call SetValue(C1a, C2)
    Set Obj = box2R1C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = True
    End With
    If F <> Empty Then F = F + " + "
    Same = Not ((CBool(Val(Left$(lblR(C1), 1)))) Xor (CBool(Val(Left$(lblR(C1a), 1)))))
    If Same Then
     F = F + Left$(lblR(0), 1)
     If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    Same = Not ((CBool(Val(Right$(lblR(C1), 1)))) Xor (CBool(Val(Right$(lblR(C1a), 1)))))
    If Same Then
     F = F + Right$(lblR(0), 1)
     If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    End If
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C2), 1) = "0" Then F = F + "' "
   End If
  Next C2
 Next C1
Mask1:
 For C1 = 1 To 4
  For C2 = 1 To 4
   If Tested(C1, C2) = "0" And txtKV(10 * C1 + C2) = "1" Then
    Call SetValue(C1, C2)
    Set Obj = box1R1C
    Load Obj(Obj.Count)
    With Obj(Obj.Count - 1)
     .Top = txtKV(10 * C1 + C2).Top - 120
     .Left = txtKV(10 * C1 + C2).Left - 120
     .Visible = True
    End With
    If F <> Empty Then F = F + " + "
    F = F + Left$(lblR(0), 1)
    If Left$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Right$(lblR(0), 1)
    If Right$(lblR(C1), 1) = "0" Then F = F + "' "
    F = F + Left$(lblC(0), 1)
    If Left$(lblC(C2), 1) = "0" Then F = F + "' "
    F = F + Right$(lblC(0), 1)
    If Right$(lblC(C2), 1) = "0" Then F = F + "' "
   Else
    Tested(C1, C2) = "1"
   End If
  Next C2
 Next C1
EndFunction:
End Function

Private Sub Command1_Click()
Dim Strings() As String

mm = FreeFile
Open "c:\circuits_3.txt" For Output As mm
Close #mm

NInp = 3
Largo = 2 ^ NInp
For Num = 0 To 2 ^ Largo - 1     'Num en binario representa una combinacion de outputs
    Me.Caption = "Num=" & Num: DoEvents
    Call PasaBinario(Num, Largo, StrBin)
    For t = 1 To Largo
        Me.txtKV(11).Text = "0"
        Me.txtKV(12).Text = "0"
        Me.txtKV(13).Text = "0"
        Me.txtKV(14).Text = "0"
        Me.txtKV(21).Text = "0"
        Me.txtKV(22).Text = "0"
        Me.txtKV(23).Text = "0"
        Me.txtKV(24).Text = "0"
        Me.txtKV(31).Text = "0"
        Me.txtKV(32).Text = "0"
        Me.txtKV(33).Text = "0"
        Me.txtKV(34).Text = "0"
        Me.txtKV(41).Text = "0"
        Me.txtKV(42).Text = "0"
        Me.txtKV(43).Text = "0"
        Me.txtKV(44).Text = "0"
        'Coloca este string en el mapa de karnaugh
        Me.txtKV(11).Text = Mid(StrBin, Largo, 1): Me.txtKV(21).Text = Mid(StrBin, Largo, 1)
        Me.txtKV(12).Text = Mid(StrBin, Largo - 1, 1): Me.txtKV(22).Text = Mid(StrBin, Largo - 1, 1)
        Me.txtKV(13).Text = Mid(StrBin, Largo - 3, 1): Me.txtKV(23).Text = Mid(StrBin, Largo - 3, 1)
        Me.txtKV(14).Text = Mid(StrBin, Largo - 2, 1): Me.txtKV(24).Text = Mid(StrBin, Largo - 2, 1)
        
        Me.txtKV(31).Text = Mid(StrBin, Largo - 4, 1): Me.txtKV(41).Text = Mid(StrBin, Largo - 4, 1)
        Me.txtKV(32).Text = Mid(StrBin, Largo - 5, 1): Me.txtKV(42).Text = Mid(StrBin, Largo - 5, 1)
        Me.txtKV(33).Text = Mid(StrBin, Largo - 7, 1): Me.txtKV(43).Text = Mid(StrBin, Largo - 7, 1)
        Me.txtKV(34).Text = Mid(StrBin, Largo - 6, 1): Me.txtKV(44).Text = Mid(StrBin, Largo - 6, 1)
    Next t
    'Presenta la funcion canonica simplificada
    Funcion = Trim(F) & " "
    
    If Trim(Funcion) <> "" Then
        numBranch = 1
        For r = 1 To Len(F)
            car = Mid(F, r, 1)
            If car = "+" Then numBranch = numBranch + 1
        Next r
        
2       'Graba los resultados
        mm = FreeFile
        Open "c:\circuits_3.txt" For Append As mm
            Print #mm, Num, numBranch
        Close #mm
    End If
    

Next Num

MsgBox "END"
End Sub

Private Sub Command2_Click()

End Sub


