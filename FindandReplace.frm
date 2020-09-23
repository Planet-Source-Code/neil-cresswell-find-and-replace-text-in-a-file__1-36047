VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Find and Replace"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Text            =   "c:\temp\test.txt"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "goodbye"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "hello"
      Top             =   840
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "(Note: Find and Replace Fields are Case Sensitive)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find and Replace Demo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "File to Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Text to Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Text to Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Declare Variables
Dim strFind As String
Dim strReplace As String
Dim strDestination As String
Dim strSource As String
Dim strFilter
Dim parse As String
Dim hold As String
Dim found As Integer
Dim num As Integer
Dim length As Integer

'Set Variables
strFind = Text1.Text 'What To find
strReplace = Text2.Text 'What To replace it With
filetosearch = Text3.Text 'File to be searched / replaced

'Open file and look for string and replace with new string

    Open filetosearch For Input As #1
    While Not EOF(1)
        Line Input #1, parse
                If InStr(parse, strFind) > 0 Then
                found = 1
                txtfile = Left(parse, InStr(parse, strFind) - 1) & strReplace & Mid(parse, InStr(parse, strFind) + Len(strFind))
                parse = txtfile
                End If

List1.AddItem parse
Wend
Close #1

'If string was replaced, write changed file back

If found = 1 Then
length = List1.ListCount
num = 0
While num < length
List1.ListIndex = num
'Write a new copy of the file
If num = "0" Then
Open filetosearch For Output As #1
Print #1, List1.Text
Close #1
Else
'Append to new file
Open filetosearch For Append As #1
Print #1, List1.Text
Close #1
End If
num = num + 1
Wend
End If
If found = 1 Then
MsgBox ("Found and Replaced at Least one instance of Search String")
Else
MsgBox ("No instanances of search string found in File")
End If
Close #1
num = 0
End Sub

