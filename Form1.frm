VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exampe of how to ""shuffle"" items in a list."
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "What is this?"
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add an item"
      Height          =   1635
      Left            =   120
      TabIndex        =   8
      Top             =   15
      Width           =   1545
      Begin VB.TextBox Text2 
         Height          =   240
         Left            =   150
         TabIndex        =   1
         Text            =   "0"
         Top             =   945
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   285
         Left            =   435
         TabIndex        =   2
         Top             =   1245
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   240
         Left            =   150
         TabIndex        =   0
         Text            =   "by G.R."
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Index:"
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Text:"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   210
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shuffle"
      Height          =   1635
      Left            =   3870
      TabIndex        =   7
      Top             =   15
      Width           =   945
      Begin VB.CommandButton Command2 
         Caption         =   "Up"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Down"
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   1065
         Width           =   645
      End
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1755
      TabIndex        =   3
      Top             =   180
      Width           =   2040
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "http://www.majiksoftware.cjb.net"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   210
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2295
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "majiksoftware@yahoo.com"
      Height          =   240
      Left            =   210
      TabIndex        =   12
      Top             =   2040
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "© 2000 by Gonzalo Ramirez"
      Height          =   240
      Left            =   210
      TabIndex        =   11
      Top             =   1785
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' | This is an exampe of how to "shuffle" items in a list.          | |
' |                                                                 | |
' | © 2000 by Gonzalo Ramirez                                       | |
' | Coded by:  Gonzalo Ramirez                                      | |
' | e-mail:    majiksoftware@yahoo.com                              | |
' | home page: http://www.majiksoftware.cjb.net                     | |
' | date:      6.19.00                                              | |
' |                                                                 | |
' | You may use this code any way you'd like (except for launching  | |
' | nuclear missles, and such). If this code has been helpful to    | |
' | you, send me an e-mail and let me know. That's all I ask. Just  | |
' | because you found this code, do not go on without learning  to  | |
' | do it youself. For that reason I have included documentation of | |
' | ever line of code.                                              | |
' |                                                  Thanks,        | |
' |                                          Gonzalo Ramirez        | |
' |                                                                 | |
' | P.S. This is my first submition to www.Planet-Source-Code.com,  | |
' |      so feedback would be nice :)                               | |
' |_________________________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\|

'Force everything to be declared.
Option Explicit

Private Sub Command1_Click()
    'Fir this example, empty items will not be added.
    If Text1.Text = "" Then Exit Sub
    
    'Make sure the index will not be out of bound
    '   and that an index was specified. If these
    '   two conditions are not met, this will set
    '   Text2.Text to List1.ListCount, which cause
    '   the item to be added at the bottom of the
    '   list.
    If (Val(Text2.Text) > List1.ListCount) Or (Text2.Text = "") Then Text2.Text = List1.ListCount

    'Add the item in the specified index.
    List1.AddItem Text1.Text, Val(Text2.Text)
    
    'Clear the text fileds
    Text1.Text = ""
    Text2.Text = ""
    
    'Set the focus back to Text1
    Text1.SetFocus
End Sub


Private Sub Command2_Click()
    'Declare variables.
    Dim tempEntry As String
    Dim tempIndex As Integer
    
    'Make sure there is something selected.
    If List1.ListIndex <> -1 Then
        'Store the text of the currently selected item in the list.
        tempEntry = List1.List(List1.ListIndex)
        'Store the index of the currently selected item in the list.
        tempIndex = List1.ListIndex
        
        'Remove the currently selected item from the list
        List1.RemoveItem List1.ListIndex
        'Add it again, but this time with an index lower than
        '   where it used to be, causeing the item on the list
        '   to seem as if it *moved* UP.
        List1.AddItem tempEntry, tempIndex - 1
        
        'Highlight the newly added item.
        List1.ListIndex = tempIndex - 1
    End If
End Sub


Private Sub Command3_Click()
    'Declare variables.
    Dim tempEntry As String
    Dim tempIndex As Integer
    
    'Make sure there is something selected.
    If List1.ListIndex <> -1 Then
        'Store the text of the currently selected item in the list.
        tempEntry = List1.List(List1.ListIndex)
        'Store the index of the currently selected item in the list.
        tempIndex = List1.ListIndex
        
        'Remove the currently selected item from the list
        List1.RemoveItem List1.ListIndex
        'Add it again, but this time with an index higher than
        '   where it used to be, causeing the item on the list
        '   to seem as if it *moved* DOWN.
        List1.AddItem tempEntry, tempIndex + 1
        
        'Highlight the newly added item.
        List1.ListIndex = tempIndex + 1
    End If
End Sub


Private Sub Command4_Click()
    Dim response As Integer
    response = MsgBox("This is an example of how to ""shuffle"" items in a list." & vbCrLf & vbCrLf & vbCrLf & "To see this code in action, add a few items to the list." & vbCrLf & vbCrLf & "After you do that, select any item from the list." & vbCrLf & vbCrLf & "You then click on UP or DOWN, depending on what" & vbCrLf & "you are allowed and on what you want to do." & vbCrLf & vbCrLf & vbCrLf & "Don't forget to read the documented comments." & vbCrLf & vbCrLf & "Thanks," & vbCrLf & "Gonzalo Ramirez", 64, "What is this?")
End Sub

Private Sub Form_Load()
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' | This is an exampe of how to "shuffle" items in a list.          | |
' |                                                                 | |
' | © 2000 by Gonzalo Ramirez                                       | |
' | Coded by:  Gonzalo Ramirez                                      | |
' | e-mail:    majiksoftware@yahoo.com                              | |
' | home page: http://www.majiksoftware.cjb.net                     | |
' | date:      6.19.00                                              | |
' |                                                                 | |
' | You may use this code any way you'd like (except for launching  | |
' | nuclear missles, and such). If this code has been helpful to    | |
' | you, send me an e-mail and let me know. That's all I ask. Just  | |
' | because you found this code, do not go on without learning  to  | |
' | do it youself. For that reason I have included documentation of | |
' | ever line of code.                                              | |
' |                                                  Thanks,        | |
' |                                          Gonzalo Ramirez        | |
' |                                                                 | |
' | P.S. This is my first submition to www.Planet-Source-Code.com,  | |
' |      so feedback would be nice :)                               | |
' |_________________________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\|

    'This is not the main focus of the example, but here it goes...
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'It is crucial that the listbox is not sorted.
    '   Please make sure that the list's "SORT" property is set to false.
    '   You can verify by clicking on the "Properties" window. From the
    '   combo dropdown menu, select "List1". Now scroll down to the
    '   property "SORT". This property can not be set at runtime.
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is not the main focus of the example, but here it goes...
    
    'Make Label5.ForeColor blue anytime the mouse is not on the label.
    Label5.ForeColor = vbBlue
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is not the main focus of the example, but here it goes...
    
    'Make Label5.ForeColor blue anytime the mouse is not on the label.
    Label5.ForeColor = vbBlue
End Sub


Private Sub Label4_DblClick()
    'This is not the main focus of the example, but here it goes...
    
    'Clear the contents of the clipboard.
    Clipboard.Clear
    'Copy the contents of Label4.Caption to the clipboard.
    Clipboard.SetText Label4.Caption
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is not the main focus of the example, but here it goes...
    
    'Make Label5.ForeColor blue anytime the mouse is not on the label.
    Label5.ForeColor = vbBlue
End Sub


Private Sub Label5_DblClick()
    'This is not the main focus of the example, but here it goes...
    
    'Use the built-in command to launch the default browser and
    '   point it to Label5.Caption, which I have set to my URL.
    Shell "start " & Label5.Caption
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is not the main focus of the example, but here it goes...
    
    'Make Label5.ForeColor red anytime the mouse is on the label.
    Label5.ForeColor = vbRed
End Sub


Private Sub List1_Click()
    'This will make sure the Command buttons are enabled/disabled
    '   properly depending on the item currently selected. The reason
    '   this is done is that if you try to shuffle UP an item already
    '   at the top, the index will be out of bound. Same goes with the
    '   item at the bottom, and for any time there is only 1 item on
    '   the list.

    If List1.ListCount = 1 Then
        'There is only 1 item on the list...Disable both Command buttons.
        Command2.Enabled = False
        Command3.Enabled = False
    ElseIf List1.ListIndex = 0 Then
        'Item selected is the first item on the list...Disable the
        '   ADD Command button and enable the DOWN Command button.
        Command2.Enabled = False
        Command3.Enabled = True
    ElseIf List1.ListIndex = List1.ListCount - 1 Then
        'Item selected is the last item on the list...Enable the
        '   ADD Command button and disable the DOWN Command button.
        Command2.Enabled = True
        Command3.Enabled = False
    Else
        'Any other item selected can be moved UP or DOWN, so enable
        '   both Command buttons.
        Command2.Enabled = True
        Command3.Enabled = True
    End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
    'Don't allow any keys to be pressed other
    '   than numbers and the backspace.
    If (KeyAscii <> 8) And (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub


