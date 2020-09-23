VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                    .: LED Matrix Display :."
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Sliding"
      Height          =   375
      Left            =   1470
      TabIndex        =   4
      Top             =   2100
      Width           =   2115
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1470
      TabIndex        =   13
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   4440
      Top             =   2100
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Right Sliding"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1740
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   1065
      Left            =   3870
      TabIndex        =   6
      Top             =   1260
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   990
      Left            =   270
      ScaleHeight     =   930
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   1290
      Width           =   915
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Text            =   "75"
         Top             =   540
         Width           =   615
      End
      Begin VB.VScrollBar vsSpeed 
         Height          =   315
         Left            =   630
         Max             =   500
         Min             =   1
         TabIndex        =   8
         Top             =   540
         Value           =   75
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Step Interval"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   645
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1065
      Left            =   240
      TabIndex        =   11
      Top             =   1260
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   3030
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Left Sliding"
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   3840
      Top             =   2100
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   3690
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Text:"
      Height          =   345
      Left            =   1470
      TabIndex        =   5
      Top             =   1230
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   10
      Left            =   4215
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   9
      Left            =   3795
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   8
      Left            =   3375
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   7
      Left            =   2955
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   6
      Left            =   2535
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   5
      Left            =   2115
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   4
      Left            =   1695
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   3
      Left            =   1275
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   2
      Left            =   855
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   1
      Left            =   435
      Top             =   285
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   240
      Top             =   90
      Width           =   4605
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim char1, char2, char3, char4, char5, char6, _
char7, char8, char9, char10 As String
Dim x As Long
Dim Char As String
Dim Str As String

Private Sub Command1_Click()
'If entered text lenght < 10 then add Spaces at the end
    Text1 = Text4
    Select Case Len(Text1)
        Case 1
            Text1 = Text1 + Space(9)
        Case 2
            Text1 = Text1 + Space(8)
        Case 3
            Text1 = Text1 + Space(7)
        Case 4
            Text1 = Text1 + Space(6)
        Case 5
            Text1 = Text1 + Space(5)
        Case 6
            Text1 = Text1 + Space(4)
        Case 7
            Text1 = Text1 + Space(3)
        Case 8
            Text1 = Text1 + Space(2)
        Case 9
            Text1 = Text1 + Space(1)
    End Select
    Text3 = Text2
    
    'Start Left Sliding
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
'Stop sliding
    Timer1.Enabled = False
    Timer2.Enabled = False
    Text1 = Text3
End Sub

Private Sub Command3_Click()
'Show About
About.Show
End Sub

Private Sub Command5_Click()
'If entered text lenght < 10 then add Spaces at the end
    Text1 = Text4

    Select Case Len(Text1)
        Case 1
            Text1 = Text1 + Space(9)
        Case 2
            Text1 = Text1 + Space(8)
        Case 3
            Text1 = Text1 + Space(7)
        Case 4
            Text1 = Text1 + Space(6)
        Case 5
            Text1 = Text1 + Space(5)
        Case 6
            Text1 = Text1 + Space(4)
        Case 7
            Text1 = Text1 + Space(3)
        Case 8
            Text1 = Text1 + Space(2)
        Case 9
            Text1 = Text1 + Space(1)
    End Select
    Text3 = Text2
    
    'Start Right Sliding
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
'Set LEDMatrix blank
x = 0
    For j = 1 To 10
        Image1(j).Picture = LoadResPicture("blank", bitmap)
    Next j
End Sub

Private Sub Text1_Change()
Text2 = Text1
Str = Text2
    
    'Blank LEDMatrix
    For j = 1 To 10
        Image1(j).Picture = LoadResPicture("blank", bitmap)
    Next j
    
    
    For i = 1 To Len(Str)
        Char = Mid(Str, i, 1)
        
        'If char= " ", "@", """, "`"
        Select Case Char
            Case " "
                Char = "blank"
            Case "@"
                Char = "at"
            Case """"
                Char = "foot"
            Case "`"
                Char = "UpKom"
        End Select
        
        'Show Char
        Image1(i).Picture = LoadResPicture(Char, bitmap)
    Next i
End Sub

Private Sub Text3_Change()
'Selecting Chars
    char1 = Mid(Text1, 1, 1)
    char2 = Mid(Text1, 2, 1)
    char3 = Mid(Text1, 3, 1)
    char4 = Mid(Text1, 4, 1)
    char5 = Mid(Text1, 5, 1)
    char6 = Mid(Text1, 6, 1)
    char7 = Mid(Text1, 7, 1)
    char8 = Mid(Text1, 8, 1)
    char9 = Mid(Text1, 9, 1)
    char10 = Mid(Text1, 10, 1)
End Sub

Private Sub Text4_Change()
Text1 = Text4
Text1_Change
End Sub

Private Sub Timer1_Timer()
    x = x + 1
    
'Left Sliding
    Select Case x
        Case 1
            Text1 = char2 + char3 + char4 + char5 + char6 _
            + char7 + char8 + char9 + char10 + char1
        Case 2
            Text1 = char3 + char4 + char5 + char6 + char7 _
            + char8 + char9 + char10 + char1 + char2
        Case 3
            Text1 = char4 + char5 + char6 + char7 + char8 _
            + char9 + char10 + char1 + char2 + char3
        Case 4
            Text1 = char5 + char6 + char7 + char8 + char9 _
            + char10 + char1 + char2 + char3 + char4
        Case 5
            Text1 = char6 + char7 + char8 + char9 + char10 _
            + char1 + char2 + char3 + char4 + char5
        Case 6
            Text1 = char7 + char8 + char9 + char10 + char1 _
            + char2 + char3 + char4 + char5 + char6
        Case 7
            Text1 = char8 + char9 + char10 + char1 + char2 _
            + char3 + char4 + char5 + char6 + char7
        Case 8
            Text1 = char9 + char10 + char1 + char2 + char3 _
            + char4 + char5 + char6 + char7 + char8
        Case 9
            Text1 = char10 + char1 + char2 + char3 + char4 _
            + char5 + char6 + char7 + char8 + char9
        Case 10
            Text1 = char1 + char2 + char3 + char4 + char5 _
            + char6 + char7 + char8 + char9 + char10

    End Select
    
    If x = 10 Then x = 0

End Sub

Private Sub Timer2_Timer()
    x = x + 1

'Right sliding
    Select Case x
        Case 1
            Text1 = char1 + char2 + char3 + char4 + char5 _
            + char6 + char7 + char8 + char9 + char10
        Case 2
            Text1 = char10 + char1 + char2 + char3 + char4 _
            + char5 + char6 + char7 + char8 + char9
        Case 3
            Text1 = char9 + char10 + char1 + char2 + char3 _
            + char4 + char5 + char6 + char7 + char8
        Case 4
            Text1 = char8 + char9 + char10 + char1 + char2 _
            + char3 + char4 + char5 + char6 + char7
        Case 5
            Text1 = char7 + char8 + char9 + char10 + char1 _
            + char2 + char3 + char4 + char5 + char6
        Case 6
            Text1 = char6 + char7 + char8 + char9 + char10 _
            + char1 + char2 + char3 + char4 + char5
        Case 7
            Text1 = char5 + char6 + char7 + char8 + char9 _
            + char10 + char1 + char2 + char3 + char4
        Case 8
            Text1 = char4 + char5 + char6 + char7 + char8 _
            + char9 + char10 + char1 + char2 + char3
        Case 9
            Text1 = char3 + char4 + char5 + char6 + char7 _
            + char8 + char9 + char10 + char1 + char2
        Case 10
            Text1 = char2 + char3 + char4 + char5 + char6 _
            + char7 + char8 + char9 + char10 + char1
    End Select
    
    
    If x = 10 Then x = 0

End Sub

Private Sub txtSpeed_Change()
'Step Interval change
vsSpeed.Value = txtSpeed.Text
Timer1.Interval = txtSpeed.Text
Timer2.Interval = txtSpeed.Text
End Sub

Private Sub vsSpeed_Change()
'Step Interval change
txtSpeed.Text = vsSpeed.Value
Timer1.Interval = txtSpeed.Text
Timer2.Interval = txtSpeed.Text
End Sub
