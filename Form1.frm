VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Allow Users to Resize Controls"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6324
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   6324
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   372
      Left            =   5400
      TabIndex        =   5
      Top             =   5160
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load more sample data"
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   2052
   End
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3972
   End
   Begin VB.Label Label3 
      Caption         =   "Example for left and right control resizing.  You can futher enhance this to resize top bottom, top-left, top-right, and etc."
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5772
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   732
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   4572
   End
   Begin VB.Label Label1 
      Caption         =   "Drag the left or right border of the list box to resize left or right!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5772
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Tip originally from Fran Pregernik, Zagreb, Croatia

Private Sub Command1_Click()
    ' Load more sample data to get a scroll bar on the list box
    ' For some reason, if you have scroll bar, it does not appear to
    ' resize from the right
    For c = 1 To 8
        List1.AddItem "This is item " & c & " as sample data."
    Next c
End Sub

Private Sub Command2_Click()
    Unload Form1
End Sub

Private Sub Form_Load()
    '  Load sample data
    For c = 1 To 8
        List1.AddItem "This is item " & c & " as sample data."
    Next c
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim nParam As Long
    
    With List1
        '  You can change these coordinates
        If (X > 0 And X < 100) Then
            nParam = HTLEFT
        ElseIf (X > .Width - 100 And X < .Width) Then
            nParam = HTRIGHT
        End If
        If nParam Then
            Call ReleaseCapture
            Call SendMessage(.hwnd, WM_NCLBUTTONDOWN, nParam, 0)
        End If
    End With
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim NewPointer As MousePointerConstants
    
    If (X > 0 And X < 100) Then
        NewPointer = vbSizeWE
    ElseIf (X > List1.Width - 100 And X < List1.Width) Then
        NewPointer = vbSizeWE
    Else
        NewPointer = vbDefault
    End If
    
    If NewPointer <> List1.MousePointer Then
        List1.MousePointer = NewPointer
    End If
    
End Sub
