VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stack Class"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Pop"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Push"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label ad 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Top Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stack Class By James Brannan"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label stop 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Top of Stack:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label it 
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items on Stack:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stack As New Stack

Private Sub Command1_Click()
Stack.Push Text1.Text
it.Caption = Stack.StackCount + 1
[stop].Caption = Stack.DebugStack(Stack.StackCount + 1)
ad.Caption = Stack.StackCount
End Sub

Private Sub Command2_Click()
'We don't have to use the If structure below, if we set ErrUnderrun to True.

If Stack.StackCount = -1 Then
    MsgBox "The stack is empty. Push a value onto it before popping.", vbInformation
Else
MsgBox "Popped " & Stack.Pop & " from the stack.", vbInformation, "Done"
it.Caption = Stack.StackCount + 1
[stop].Caption = Stack.DebugStack(Stack.StackCount + 1)
ad.Caption = Stack.StackCount
End If
End Sub

Private Sub Form_Load()
Set Stack = New Stack
Stack.ErrUnderrun = False
'We're going to use our own error handling
Stack.StkLimit = -1
'Allow unlimited items on stack

it.Caption = Stack.StackCount + 1
[stop].Caption = "-1"
ad.Caption = "N/A"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Stack = Nothing
End Sub

