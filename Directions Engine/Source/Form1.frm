VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Directions"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   ForeColor       =   &H80000006&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox IntroPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   7815
      ScaleWidth      =   11895
      TabIndex        =   12
      Top             =   0
      Width           =   11895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   0
      Picture         =   "Form1.frx":3089
      ScaleHeight     =   1342.268
      ScaleMode       =   0  'User
      ScaleWidth      =   12015
      TabIndex        =   11
      Top             =   0
      Width           =   12015
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   3480
      Picture         =   "Form1.frx":3F13
      ScaleHeight     =   5985
      ScaleWidth      =   4905
      TabIndex        =   8
      Top             =   1800
      Width           =   4935
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   0
         Picture         =   "Form1.frx":67CC
         ScaleHeight     =   6255
         ScaleWidth      =   4935
         TabIndex        =   10
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label NewspaperTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      Picture         =   "Form1.frx":9085
      ScaleHeight     =   2415
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   1560
      Width           =   12015
      Begin VB.Label Question 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   11055
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Option6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   7200
      Width           =   7335
   End
   Begin VB.Label Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   6600
      Width           =   7335
   End
   Begin VB.Label Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   6000
      Width           =   7335
   End
   Begin VB.Label Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   5400
      Width           =   7335
   End
   Begin VB.Label Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   7335
   End
   Begin VB.Label Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   7335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadItem 
         Caption         =   "Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu munPlayItem 
         Caption         =   "Play"
         Shortcut        =   ^P
      End
      Begin VB.Menu munQuitItem 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
            
'*-------------*
'*             *
'* Key Checker *
'* ~~~~~~~~~~~ *
'*             *
'*-------------*
    
    Select Case KeyAscii
    
        Case KEY_SPACE, KEY_ENTER
            
            '*-------------------------------------------------*
            '* enter or space pressed, setup next options page *
            '*-------------------------------------------------*
            
            If StartGame And ScenarioLoaded Then
            
                '**** Game has just started
                
                IntroPicture.Visible = False
                StartGame = False
                
                '**** Check for initial headline
    
                Picture3_Click
    
                '**** run game
    
                GameLoop
            
            Else
                
                If OptionPage = False And ScenarioLoaded Then

                    Picture2.Visible = True
                    Question.Visible = True
                    
                    Picture3.Visible = False
                    Picture4.Visible = False
                    NewspaperTitle.Visible = False
                
                    GameLoop
            
                End If
            
            End If
 
        Case KEY_1
        
            '*-----------------*
            '*                 *
            '* Option 1 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option1.Visible And StartGame = False Then
            
                Call ChooseOption(1, CurrentNode)
            
            End If
    
        Case KEY_2
        
            '*-----------------*
            '*                 *
            '* Option 2 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option2.Visible And StartGame = False Then
            
                Call ChooseOption(2, CurrentNode)
            
            End If
            
        Case KEY_3
        
            '*-----------------*
            '*                 *
            '* Option 3 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option3.Visible And StartGame = False Then
            
                Call ChooseOption(3, CurrentNode)
            
            End If
            
        Case KEY_4
        
            '*-----------------*
            '*                 *
            '* Option 4 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option4.Visible And StartGame = False Then
            
                Call ChooseOption(4, CurrentNode)
            
            End If
            
        Case KEY_5
        
            '*-----------------*
            '*                 *
            '* Option 5 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option5.Visible And StartGame = False Then
            
                Call ChooseOption(5, CurrentNode)
            
            End If
            
        Case KEY_6
        
            '*-----------------*
            '*                 *
            '* Option 6 chosen *
            '* ~~~~~~~~~~~~~~~ *
            '*                 *
            '*-----------------*
            
            If OptionPage And Option6.Visible And StartGame = False Then
            
                Call ChooseOption(6, CurrentNode)
            
            End If
            
        Case KEY_P
        
            Picture3_Click
   
    End Select
 
End Sub

Private Sub Form_Load()

    '*-------*
    '*       *
    '* Setup *
    '* ~~~~~ *
    '*       *
    '*-------*
    
    ScenarioLoaded = False
    StartGame = True
    
    '**** setup background color
    
    Form1.BackColor = RGB(FORM_COLOR_RED, FORM_COLOR_GREEN, FORM_COLOR_BLUE)
    frmAbout.BackColor = RGB(FORM_COLOR_RED, FORM_COLOR_GREEN, FORM_COLOR_BLUE)
    Picture2.BackColor = RGB(FORM_COLOR_RED, FORM_COLOR_GREEN, FORM_COLOR_BLUE)
    
    '**** Set control visibility
    
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    NewspaperTitle.Visible = False
    
    IntroPicture.Height = Form1.Height
    IntroPicture.Width = Form1.Width
    
    KeyPreview = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = &H80000006
    Option3.ForeColor = &H80000006
    Option4.ForeColor = &H80000006
    Option5.ForeColor = &H80000006
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Form_Terminate()
    
    '*------*
    '*      *
    '* Quit *
    '* ~~~~ *
    '*      *
    '*------*
     
    Unload Form1
    End

End Sub

Private Sub Option1_Click()

    '*-----------------*
    '*                 *
    '* Option 1 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(1, CurrentNode)
         
End Sub

Private Sub Option2_Click()

    '*-----------------*
    '*                 *
    '* Option 2 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(2, CurrentNode)
         
End Sub

Private Sub Option3_Click()

    '*-----------------*
    '*                 *
    '* Option 3 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(3, CurrentNode)
         
End Sub

Private Sub Option4_Click()

    '*-----------------*
    '*                 *
    '* Option 4 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(4, CurrentNode)
         
End Sub

Private Sub Option5_Click()

    '*-----------------*
    '*                 *
    '* Option 5 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(5, CurrentNode)
         
End Sub

Private Sub Option6_Click()

    '*-----------------*
    '*                 *
    '* Option 6 chosen *
    '* ~~~~~~~~~~~~~~~ *
    '*                 *
    '*-----------------*
    
     Call ChooseOption(6, CurrentNode)
         
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    Option2.ForeColor = &H80000006
    Option3.ForeColor = &H80000006
    Option4.ForeColor = &H80000006
    Option5.ForeColor = &H80000006
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    Option3.ForeColor = &H80000006
    Option4.ForeColor = &H80000006
    Option5.ForeColor = &H80000006
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = &H80000006
    Option3.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    Option4.ForeColor = &H80000006
    Option5.ForeColor = &H80000006
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Option4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = &H80000006
    Option3.ForeColor = &H80000006
    Option4.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    Option5.ForeColor = &H80000006
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Option5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = &H80000006
    Option3.ForeColor = &H80000006
    Option4.ForeColor = &H80000006
    Option5.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    Option6.ForeColor = &H80000006
    
End Sub

Private Sub Option6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Option1.ForeColor = &H80000006
    Option2.ForeColor = &H80000006
    Option3.ForeColor = &H80000006
    Option4.ForeColor = &H80000006
    Option5.ForeColor = &H80000006
    Option6.ForeColor = RGB(HOVER_COLOR_RED, HOVER_COLOR_GREEN, HOVER_COLOR_BLUE)
    
End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.BackColor = RGB(FORM_COLOR_RED, FORM_COLOR_GREEN, FORM_COLOR_BLUE)
    frmAbout.Visible = True

End Sub

Private Sub mnuLoadItem_Click()
    
    '*------*
    '*      *
    '* Load *
    '* ~~~~ *
    '*      *
    '*------*
    
    '*-------------------------------*
    '* Open Dialog and load filename *
    '*-------------------------------*
    
    '**** Trap File Error
    
    On Error GoTo Error_Handler
    
    '**** Open Dialog box
    
    CommonDialog1.ShowOpen
 
    '**** Open And process file
 
    Open CommonDialog1.FileName For Binary As #1
    
    Get #1, 1, NumberOfNodes
    
    ReDim IsAction(NumberOfNodes)
    ReDim Description(NumberOfNodes)
    ReDim Headline(NumberOfNodes)
    ReDim ChanceOfSuccess(NumberOfNodes)
    ReDim Perception(NumberOfNodes)
    ReDim Parent(NumberOfNodes)
    ReDim NodeTitles(NumberOfNodes)
    ReDim PowerPointSlidePaths(NumberOfNodes)
    
    Get #1, , IsAction
    Get #1, , Description
    Get #1, , Headline
    Get #1, , ChanceOfSuccess
    Get #1, , Perception
    Get #1, , Parent
    Get #1, , NodeTitles
    Get #1, , PowerPointSlidePaths
    
    Close #1
  
    ScenarioLoaded = True
    
    '**** Start game
    
    CurrentNode = 1
    
    '**** Show/hide controls
     
    Picture2.Visible = True
    Picture3.Visible = False
    Question.Visible = True
    
    IntroPicture.Visible = True
    StartGame = True
    Exit Sub
    
Error_Handler:
    
    '**** No node is selected, show error message box
    
    Dim result As VbMsgBoxResult
    result = MsgBox("Invalid Filename", , "File Error")
        
End Sub

Private Sub munPlayItem_Click()
    
    '*------*
    '*      *
    '* Play *
    '* ~~~~ *
    '*      *
    '*------*
     
    If Not ScenarioLoaded Then
    
        Dim result As VbMsgBoxResult
        result = MsgBox("No Scenario Loaded", , "Directions Error")
    
    Else
    
        '**** Start game
        
        CurrentNode = 1
        
        '**** Show/hide controls
         
        Picture2.Visible = True
        Picture3.Visible = False
        Question.Visible = True
    
        IntroPicture.Visible = True
        StartGame = True
    
    End If
    
End Sub

Private Sub munQuitItem_Click()
    
    '*------*
    '*      *
    '* Quit *
    '* ~~~~ *
    '*      *
    '*------*
     
    Unload Form1
    End
    
End Sub

Private Function PopulateQuestionAndOptions(CurrentActionNode As Integer) As String

    '*-------------------------------*
    '*                               *
    '* Populate Question and Options *
    '* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ *
    '*                               *
    '*-------------------------------*
      
    Dim NodeCounter As Integer
    NodeCounter = 0
    
    '**** Options
    
    OptionPage = True

    '**** Disable Options
    
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    Option5.Visible = False
    Option6.Visible = False
    
    '**** Fill Question
    
    Question.Caption = Description(CurrentActionNode)
    
    '*------------------------------------------------*
    '* iterate through, populate Question and options *
    '*------------------------------------------------*
     
    For i = 1 To NumberOfNodes
     
        '*-------------------------------------------------*
        '* If parent of node is current node, then process *
        '*-------------------------------------------------*
        
        If Parent(i) = CurrentActionNode Then
        
            NodeCounter = NodeCounter + 1
            
            Select Case (NodeCounter)
            
                Case 1
                    
                    Option1.Visible = True
                    Option1.Caption = "1." & Description(i)
                    
                Case 2
                
                    Option2.Visible = True
                    Option2.Caption = "2." & Description(i)
                    
                Case 3
                     
                    Option3.Visible = True
                    Option3.Caption = "3." & Description(i)
                    
                Case 4
                
                    Option4.Visible = True
                    Option4.Caption = "4." & Description(i)
                    
                Case 5
                
                    Option5.Visible = True
                    Option5.Caption = "5." & Description(i)
                    
                Case 6
                
                    Option6.Visible = True
                    Option6.Caption = "6." & Description(i)
                                    
            End Select
        
        End If
     
    Next
     
    '*-----------------------------------------------*
    '* Return appropriate value (END or NORMAL node) *
    '*-----------------------------------------------*
    
    If NodeCounter = 0 Then
    
        PopulateQuestionAndOptions = "END"
    
    Else
    
        PopulateQuestionAndOptions = "NORMAL"
    
    End If
     
End Function

Private Sub GameLoop()

    '*----------------*
    '*                *
    '* Main Game Loop *
    '* ~~~~~~~~~~~~~~ *
    '*                *
    '*----------------*
    
    Dim NodeType As String
    
    NodeType = PopulateQuestionAndOptions(CurrentNode)
    
End Sub

Private Sub ChooseOption(OptionNodeNumber As Integer, CurrentActionNode As Integer)

    '*-----------------------------*
    '*                             *
    '* Randomize and choose action *
    '* ~~~~~~~~~~~~~~~~~~~~~~~~~~~ *
    '*                             *
    '*-----------------------------*
    
    Dim NewActionNode As Integer
    Dim RandomValue As Integer
    Dim RandomTotal As Integer
    Dim OptionNodeIndex As Integer
    Dim NodeCounter As Integer
    NodeCounter = 0
   
    '*------------------------------------------*
    '* Iterate through, find real option number *
    '*------------------------------------------*
     
    For i = 1 To NumberOfNodes
         
        '*-----------------------------------------------------------------------------*
        '* if parent is current action, find inc. nodenumber until node index is found *
        '*-----------------------------------------------------------------------------*
         
        If Parent(i) = CurrentActionNode Then
        
            NodeCounter = NodeCounter + 1
            
            If NodeCounter = OptionNodeNumber Then
            
                OptionNodeIndex = i
            
            End If
        
        End If
         
    Next
        
    '**** Reset node counter for random divider
        
    NodeCounter = 0
        
    '*-------------------------*
    '* Calculate random total  *
    '*-------------------------*
    
    For i = 1 To NumberOfNodes
         
        '*--------------------------------------------------------------------------------------*
        '* If parent is current option, find inc. nodenumber until node Actions index are found *
        '*--------------------------------------------------------------------------------------*
         
        If Parent(i) = OptionNodeIndex Then
        
            NodeCounter = NodeCounter + 1
        
            RandomTotal = RandomTotal + (ChanceOfSuccess(i) + 1)
        
        End If
         
    Next
        
    '*------------------------*
    '* Generate random number *
    '*------------------------*
    
    Randomize
    
    RandomValue = Int(RandomTotal * Rnd + 1)
             
    '*----------------------*
    '* Find new action node *
    '*----------------------*
         
    NodeCounter = 0
      
    For i = 1 To NumberOfNodes
         
        '*--------------------------------------------------------------------------------------*
        '* If parent is current option, find inc. nodenumber until node Actions index are found *
        '*--------------------------------------------------------------------------------------*
         
        If Parent(i) = OptionNodeIndex Then
          
            If RandomValue <= (ChanceOfSuccess(i) + 1) Then
            
                NewActionNode = i
            
                Exit For
            
            End If
        
            RandomValue = RandomValue - (ChanceOfSuccess(i) + 1)
        
        End If
         
    Next
                
    '*----------------------------------*
    '* Assign current node and continue *
    '*----------------------------------*
    
    CurrentNode = NewActionNode
    
    '*------------------------*
    '* set up headline screen *
    '*------------------------*
    
    Picture2.Visible = False
    Question.Visible = False
    
    Option1.Visible = False
    Option2.Visible = False
    Option3.Visible = False
    Option4.Visible = False
    Option5.Visible = False
    Option6.Visible = False
   
    '**** Check for good/bad perception
    
    If Perception(CurrentNode) = 1 Then
    
        Picture3.Visible = True
        Picture4.Visible = False
    
    Else
    
        Picture3.Visible = False
        Picture4.Visible = True
        
    End If
    
    NewspaperTitle.Visible = True
    NewspaperTitle.Caption = Headline(CurrentNode)
    
    '**** Not options
    
    OptionPage = False

End Sub

Private Sub Picture3_Click()

    '*----------------------------------------------------------*
    '* If the headline text is not empty, attempt to show slide *
    '*----------------------------------------------------------*

    '**** PP stuff to come
    
End Sub

Private Sub Picture4_Click()

    Picture3_Click
    
End Sub
