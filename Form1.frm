VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPeg Header Info"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstExif 
      Height          =   5325
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Double Click to change"
      Top             =   120
      Width           =   5055
   End
   Begin VB.FileListBox FileList 
      Height          =   5355
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.DirListBox Directory 
      Height          =   5040
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Now downloaded "Exif2-2.PDF" format standard, April 2002 + Jap addon pack (its in english), somewhere like www.exif.org ?
' This program does not use all the features of V2.2, not even V2.1 (1999)

' Intel and Motorol Bytes
' E.G. &H123456 would be save as Motorol 12 34 56, Intel saves it as 56 34 12 ! annoying but good for the CopyMemory function

' Think of Exif as a partition table, wheres there is directories and offsets to values

' Familiarize your self with the smaller functions before you read 'ReadExif'

Option Explicit

Private Sub Directory_Change()
    FileList = Directory
End Sub

Private Sub Drive_Change()
    Directory = Drive
End Sub

Private Sub FileList_Click()
    Dim ret As Boolean
    
    ret = ReadExif(Directory & "\" & FileList, lstExif)
    'If ret = False Then MsgBox "There was an error processing a perfect file"
End Sub

Private Sub Form_Load()
    frmMain.Show
    FileList.ListIndex = 0
End Sub

Private Sub lstExif_DblClick()
    frmEdit.StartIt lstExif.ListIndex, Left$(lstExif, InStr(1, lstExif, "=") - 2), Mid$(lstExif, InStr(1, lstExif, "=") + 2)
End Sub
