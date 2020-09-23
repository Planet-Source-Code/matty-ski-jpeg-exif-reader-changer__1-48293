VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Value"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdWarn 
      Caption         =   "&Warning"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAct 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtDispVal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.269736842105263"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtNewFract 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "456"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtNewFract 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "123"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txtNewVal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblCurrVal 
      Caption         =   "X"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Display Value:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "New 'Rational' Values   ( 1 / 2 ):"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblTypeofVal 
      Caption         =   "X"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Type of Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "New Value 2:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "New Value 1:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "New Value:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Current Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblTagName 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a new value for:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheListIndex As Long

Function StartIt(LstIndex As Long, TagName As String, CurrVal As String)
    lblTagName.Caption = TagName
    lblCurrVal.Caption = CurrVal
    TheListIndex = LstIndex
    
    Select Case CLng("&H" & ListInfo(1, LstIndex))
    Case 1: lblTypeofVal = "Unsigned Byte": ShowBoxes 0
    Case 2: lblTypeofVal = "ASCII Strings": ShowBoxes 0
    Case 3: lblTypeofVal = "Unsigned Short": ShowBoxes 0
    Case 4: lblTypeofVal = "Unsigned Long": ShowBoxes 0
    Case 5: lblTypeofVal = "Unsigned Rational": ShowBoxes 1
    Case 6: lblTypeofVal = "Signed Byte": ShowBoxes 0
    Case 7: lblTypeofVal = "Undefined": ShowBoxes 0
    Case 8: lblTypeofVal = "Signed Short": ShowBoxes 0
    Case 9: lblTypeofVal = "Signed Long": ShowBoxes 0
    Case 10: lblTypeofVal = "Signed Rational": ShowBoxes 1
    Case 11: lblTypeofVal = "Single Float": ShowBoxes 0
    Case 12: lblTypeofVal = "Double Float": ShowBoxes 0
    Case Else: Beep: Unload Me: Exit Function
    End Select
    
    Me.Show , frmMain
    DoEvents
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Function

Private Sub cmdAct_Click(Index As Integer)
    Dim tmpVal As String
    
    If Index = 0 Then
        ' Confirm its a numeric value entered
        If InStr(1, lblTypeofVal.Caption, "Byte") Or InStr(1, lblTypeofVal.Caption, "Short") Or InStr(1, lblTypeofVal.Caption, "Long") Then
            If Not IsNumeric(txtNewVal) Then MsgBox "New Value must be numeric", vbExclamation, "Error": Exit Sub
        End If
        
        If txtNewVal.Enabled = True Then tmpVal = txtNewVal Else tmpVal = txtNewFract(0)
        Call ChangeExif(frmMain.Directory & "\" & frmMain.FileList, TheListIndex, tmpVal, txtNewFract(1))
    End If
    Unload Me
End Sub

Private Sub cmdWarn_Click()
    MsgBox "Ajusting some values could corrupt the file," & vbCr & "like increasing the length of a string value past its original length or changing 'undefined' types" & vbCr & vbCr & "Some values have been interpreted from there number value to a string" & vbCr & "And there is no proper error/format checking", vbExclamation, "Warning"
End Sub

Private Sub txtNewFract_Change(Index As Integer)
    On Error Resume Next
    
    txtDispVal.Text = Val(txtNewFract(0).Text) / (txtNewFract(1).Text)
    If Err Then txtDispVal.Text = Err.Description
End Sub

Private Sub txtNewFract_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 8 And Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 1
End Sub

Private Sub ShowBoxes(Boxes2Show As Byte)
    If Boxes2Show = 0 Then
        ' Display single box
        Label3.Enabled = True
        txtNewVal.Enabled = True
        txtNewVal.Text = lblCurrVal.Caption
    Else
        ' Display 2 bottom boxes for a kinda fraction
        Label8.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = True
        Label9.Enabled = True
        txtNewFract(0).Enabled = True
        txtNewFract(1).Enabled = True
        txtDispVal.Enabled = True
    End If
End Sub
