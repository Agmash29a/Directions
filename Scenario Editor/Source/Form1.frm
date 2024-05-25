VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Directions - Scenario Editor"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8705
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      MouseIcon       =   "Form1.frx":0442
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action Data"
      Height          =   1815
      Left            =   6120
      TabIndex        =   7
      Top             =   5040
      Width           =   4815
      Begin VB.ComboBox PerceptionType 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox ChanceOfHappening 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Text            =   "ChanceOfHappening"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox HeadlineText 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox PowerPointSlidePath 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Silde Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Perception:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Chance:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Headline:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "blah"
      Filter          =   "*.dlv"
   End
   Begin VB.Frame Frame6 
      Caption         =   "Nodes"
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
      Begin VB.CommandButton Update 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton DeleteNode 
         Caption         =   "Delete Node"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00808080&
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton AddChildNode 
         Caption         =   "Add Child Node"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option/Question Text"
      Height          =   1815
      Left            =   1680
      TabIndex        =   1
      Top             =   5040
      Width           =   4335
      Begin VB.TextBox DescriptionText 
         Height          =   1485
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   15
         Left            =   1440
         TabIndex        =   2
         Top             =   2040
         Width           =   135
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadItem 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuSaveItem 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuQuitItem 
         Caption         =   "Quit"
      End
      Begin VB.Menu mnuLineItem 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuScenario 
      Caption         =   "Scenario"
      Begin VB.Menu mnuParseScenarioItem 
         Caption         =   "Parse Scenario"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddChildNode_Click()
    
    '*----------------------*
    '*                      *
    '* Add a new child node *
    '* ~~~~~~~~~~~~~~~~~~~~ *
    '*                      *
    '*----------------------*
    
    Dim nd As Node
    Dim oa As OptionAction
    Dim keytext As String
    
    '*------------------------------------------*
    '* First Check to see if a node is selected *
    '*------------------------------------------*
    
    If TreeView1.SelectedItem Is Nothing Then
        
        '**** No node is selected, show error message box
    
        Dim result As VbMsgBoxResult
        result = MsgBox("You must select a node", , "Editor Error")
    
    Else

        '**** A Node is selected, add a node and grab a new key

        keytext = "key" & GetNewKey()
        
        '*------------------------*
        '* Check Parent Node Type *
        '*------------------------*
        
        If OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).IsAction Then
        
            '**** Parent Node is an Action
            
            '*------------------------------------------------------------*
            '* count number of siblings, err msg if over MAX_OPTION_NODES *
            '*------------------------------------------------------------*
            
            If TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Children < MAX_OPTION_NODES Then
            
                '**** Spare siblings left - add new Nodes
            
                Set oa = OptionsAndActions.Add(False, "", "", CHANCE_AVERAGE, PERCEPTION_TYPE_GOOD, "", keytext)
                Set nd = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, keytext, "Option", PICTURE_OPTION_NODE)
   
            Else
            
                '**** Max Options exceeded, show error message box
                
                Dim result2 As VbMsgBoxResult
                result2 = MsgBox("Only " & Str(MAX_OPTION_NODES) & " Options allowed", , "Editor Error")
        
            End If
            
        Else
        
            '**** Node is an action, create nodes
        
            Set oa = OptionsAndActions.Add(True, "", "", CHANCE_AVERAGE, PERCEPTION_TYPE_GOOD, "", keytext)
            Set nd = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, keytext, "Action", PICTURE_ACTION_NODE)
            
        End If
        
        '**** Clean up
        
        Set nd = Nothing
        Set oa = Nothing
    
    End If

End Sub

Private Sub DeleteNode_Click()
    
    '*---------------*
    '*               *
    '* Delete a node *
    '* ~~~~~~~~~~~~~ *
    '*               *
    '*---------------*
    
    '*------------------------------------------*
    '* First Check to see if a node is selected *
    '*------------------------------------------*
    
    If TreeView1.SelectedItem Is Nothing Then
    
        '**** No node is selected, show error message box
    
        Dim result As VbMsgBoxResult
        result = MsgBox("You must select a node", , "Editor Error")
    
    Else
    
        '**** A Node is selected
    
        '*-----------------------------------------*
        '* Check to see if a root node is selected *
        '*-----------------------------------------*

        If TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Parent Is Nothing Then
        
            '**** A root Node is selected, show error message box
    
            Dim result2 As VbMsgBoxResult
            result = MsgBox("You cannot delete this node", , "Editor Error")
    
           
        Else
        
            '**** Delete this node
        
            TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
        
        End If
    
    End If
    
End Sub

Private Sub Form_Load()

    '*-------*
    '*       *
    '* Setup *
    '* ~~~~~ *
    '*       *
    '*-------*
    
    Dim nd As Node
    Dim oa As OptionAction
    Dim keytext As String
    
    '*-------------------*
    '* Setup all Globals *
    '*-------------------*
    
    KeyCounter = 0
    
    PerceptionType.AddItem ("Good")
    PerceptionType.AddItem ("Bad")
    
    PerceptionType.ListIndex = PERCEPTION_TYPE_GOOD
    
    ChanceOfHappening.AddItem ("Very High")
    ChanceOfHappening.AddItem ("High")
    ChanceOfHappening.AddItem ("Average")
    ChanceOfHappening.AddItem ("Low")
    ChanceOfHappening.AddItem ("Very Low")
    
    ChanceOfHappening.ListIndex = CHANCE_AVERAGE

    '*-----------------*
    '* Setup root Node *
    '*-----------------*

    keytext = "key" & GetNewKey()
        
    Set oa = OptionsAndActions.Add(True, "", "", CHANCE_AVERAGE, PERCEPTION_TYPE_GOOD, "", keytext)
    Set nd = TreeView1.Nodes.Add(, , keytext, "Game", PICTURE_GAME_NODE)
        
    '**** Clean up
        
    Set nd = Nothing
    Set oa = Nothing
        
End Sub

Private Sub Form_Resize()

    '*--------*
    '*        *
    '* Resize *
    '* ~~~~~~ *
    '*        *
    '*--------*
    
    '**** set width of TV control
    
    TreeView1.Width = ScaleWidth - TreeView1.Left
    
    '*------------------------------------------------------------------------------*
    '* Set heights of TV and frames if scaleheight is greater than a certain amount *
    '*------------------------------------------------------------------------------*
    
    If ScaleHeight > (HEIGHT_MINIMUM_RESIZE) Then
        
        TreeView1.Height = (ScaleHeight - HEIGHT_NON_TREEVIEW) - TreeView1.Top
        
        Frame1.Top = HEIGHT_ORIGINAL_FRAMES - (HEIGHT_ORIGINAL_TREEVIEW - TreeView1.Height)
        Frame6.Top = HEIGHT_ORIGINAL_FRAMES - (HEIGHT_ORIGINAL_TREEVIEW - TreeView1.Height)
        Frame3.Top = HEIGHT_ORIGINAL_FRAMES - (HEIGHT_ORIGINAL_TREEVIEW - TreeView1.Height)
    
    End If
    
    '*----------------------------------------------------------------------------------------*
    '* Set widths of frames and internal boxes if scalewidth is greater than a certain amount *
    '*----------------------------------------------------------------------------------------*
    
    If ScaleWidth > WIDTH_ORIGINAL_FORM Then
    
        '**** set new width (frame1) and move rightmost frame right
    
        Frame1.Width = WIDTH_ORIGINAL_FRAME1 + (ScaleWidth - WIDTH_ORIGINAL_FORM)
        
        Frame3.Left = LEFT_ORIGINAL_FRAME3 + (ScaleWidth - WIDTH_ORIGINAL_FORM)
        DescriptionText.Width = WIDTH_ORIGINAL_DESCRIPTION_TEXT + (ScaleWidth - WIDTH_ORIGINAL_FORM)
    Else
    
        '**** Reset Original Values
    
        Frame1.Width = WIDTH_ORIGINAL_FRAME1
        Frame3.Left = LEFT_ORIGINAL_FRAME3
        DescriptionText.Width = WIDTH_ORIGINAL_DESCRIPTION_TEXT
        
    End If
    
End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.Visible = True

End Sub

Private Sub mnuLoadItem_Click()
    
    '*------*
    '*      *
    '* Load *
    '* ~~~~ *
    '*      *
    '*------*
     
    Dim IsAction() As Boolean
    Dim Description() As String
    Dim Headline() As String
    Dim ChanceOfSuccess() As Integer
    Dim Perception() As Integer
    Dim Parent() As Integer
    Dim NodeTitles() As String
    Dim PowerPointSlidePaths() As String
    Dim NumberOfNodes As Long
    Dim RootNodeQuestion As String
    
    Dim nd As Node
    Dim oa As OptionAction
    Dim keytext As String
    
    '*----------------------------------*
    '* Reset Collections and keycounter *
    '*----------------------------------*
    
    OptionsAndActions.RemoveAllItems
    TreeView1.Nodes.Clear
    KeyCounter = 0

    '**** Clean up
         
    Set nd = Nothing
    Set oa = Nothing
          
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
    
    Call TreeViewBuild(TreeView1, IsAction, Description, Headline, ChanceOfSuccess, Perception, Parent, NodeTitles, PowerPointSlidePaths, 0)
    
    Exit Sub
    
Error_Handler:
    
    '**** No node is selected, show error message box
    
    Dim result As VbMsgBoxResult
    result = MsgBox("Invalid Filename", , "File Error")
        
End Sub

Private Sub mnuParseScenarioItem_Click()
    
    '*----------------*
    '*                *
    '* Parse Scenario *
    '* ~~~~~~~~~~~~~~ *
    '*                *
    '*----------------*
     
    Errors = False

    '**** Call Parsing routine

    Call TreeViewParse(TreeView1, OptionsAndActions, , True)

    '*----------------------------------------*
    '* Show Msg box with Success or fail text *
    '*----------------------------------------*

    If Errors Then
    
        Dim result As VbMsgBoxResult
        result = MsgBox("Parser Error", , "Parser")
    
    Else
    
        Dim result2 As VbMsgBoxResult
        result2 = MsgBox("Success!", , "Parser")
    
    End If
    
End Sub

Private Sub mnuQuitItem_Click()
    
    '*------*
    '*      *
    '* Quit *
    '* ~~~~ *
    '*      *
    '*------*
     
    Unload Form1
    End
    
End Sub

Private Sub mnuSaveItem_Click()
    
    '*------*
    '*      *
    '* Save *
    '* ~~~~ *
    '*      *
    '*------*
     
    Dim IsAction() As Boolean
    Dim Description() As String
    Dim Headline() As String
    Dim ChanceOfSuccess() As Integer
    Dim Perception() As Integer
    Dim Parent() As Integer
    Dim NodeTitles() As String
    Dim PowerPointSlidePaths() As String
        
    '*----------------------------------*
    '* Dimension arrays to no. of nodes *
    '*----------------------------------*
        
    ReDim IsAction(TreeView1.Nodes.Count)
    ReDim Description(TreeView1.Nodes.Count)
    ReDim Headline(TreeView1.Nodes.Count)
    ReDim ChanceOfSuccess(TreeView1.Nodes.Count)
    ReDim Perception(TreeView1.Nodes.Count)
    ReDim Parent(TreeView1.Nodes.Count)
    ReDim NodeTitles(TreeView1.Nodes.Count)
    ReDim PowerPointSlidePaths(TreeView1.Nodes.Count)
    
    Dim nd As Node
       
    '*--------------------------------------------*
    '* Iterate through collections ad fill arrays *
    '*--------------------------------------------*
          
    For Each nd In TreeView1.Nodes
           
        '*--------------------------------------*
        '* Check each node for root, set parent *
        '*--------------------------------------*
        
        If Not nd.Parent Is Nothing Then
        
            Parent(nd.Index) = nd.Parent.Index
        
        End If
         
        '**** Fill Other Arrays with values
        
        NodeTitles(nd.Index) = nd.Text
        IsAction(nd.Index) = OptionsAndActions.Item(nd.Key).IsAction
        Description(nd.Index) = OptionsAndActions.Item(nd.Key).Description
        Headline(nd.Index) = OptionsAndActions.Item(nd.Key).Headline
        ChanceOfSuccess(nd.Index) = OptionsAndActions.Item(nd.Key).ChanceOfSuccess
        Perception(nd.Index) = OptionsAndActions.Item(nd.Key).Perception
        PowerPointSlidePaths(nd.Index) = OptionsAndActions.Item(nd.Key).PowerPointSlidePath
    
    Next
          
    '*-------------------------------*
    '* Save Dialog and load filename *
    '*-------------------------------*
    
    '**** Trap File Error
    
    On Error GoTo Error_Handler
    
    '**** Save Dialog box
        
    CommonDialog1.ShowSave

    '**** Open And process file
 
    Dim Intfornumnodes As Integer
     
    Open CommonDialog1.FileName For Binary As #1
    
    Put #1, 1, CLng(TreeView1.Nodes.Count)
    Put #1, , IsAction
    Put #1, , Description
    Put #1, , Headline
    Put #1, , ChanceOfSuccess
    Put #1, , Perception
    Put #1, , Parent
    Put #1, , NodeTitles
    Put #1, , PowerPointSlidePaths
 
    Close #1
    
    '**** Clean up
         
    Set nd = Nothing
    
    Exit Sub
    
Error_Handler:
    
    '**** No node is selected, show error message box
    
    Dim result As VbMsgBoxResult
    result = MsgBox("Invalid Filename", , "File Error")
    
    '**** Clean up
         
    Set nd = Nothing
                
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '*-------------------------------------*
    '*                                     *
    '* Fill Description and enable/disable *
    '* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ *
    '*                                     *
    '*-------------------------------------*
     
    '**** Fill description text
    
    DescriptionText.Text = OptionsAndActions.Item(Node.Key).Description
       
    '*--------------------------------------------*
    '* Enable/Disable (special check for root(1)) *
    '*--------------------------------------------*
    
    If (OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).IsAction = False) Or (TreeView1.SelectedItem.Index = 1) Then

        '**** Option or root node
        
        PerceptionType.Enabled = False
        ChanceOfHappening.Enabled = False
        HeadlineText.Enabled = False
        PowerPointSlidePath.Enabled = False
        
        '**** Special case for root node
        
        If TreeView1.SelectedItem.Index = 1 Then
        
            PowerPointSlidePath.Enabled = True
        
        End If
        
    Else
    
        '**** Action
    
        PerceptionType.Enabled = True
        ChanceOfHappening.Enabled = True
        HeadlineText.Enabled = True
        PowerPointSlidePath.Enabled = True
        
        '**** Fill fields
        
        PerceptionType.ListIndex = OptionsAndActions.Item(Node.Key).Perception
        ChanceOfHappening.ListIndex = OptionsAndActions.Item(Node.Key).ChanceOfSuccess
        HeadlineText.Text = OptionsAndActions.Item(Node.Key).Headline
        PowerPointSlidePath.Text = OptionsAndActions.Item(Node.Key).PowerPointSlidePath
    
    End If
    
End Sub

Private Sub Update_Click()
    
    '*--------------------*
    '*                    *
    '* Fill OA collection *
    '* ~~~~~~~~~~~~~~~~~~ *
    '*                    *
    '*--------------------*
         
    '*-----------------------------*
    '* Check if a node is selected *
    '*-----------------------------*
         
    If TreeView1.SelectedItem Is Nothing Then
   
        '**** No node selected, error msg
        
        Dim result As VbMsgBoxResult
        result = MsgBox("You must select a node", , "Editor Error")
    
    Else
    
        '**** Fill OA fields
    
        OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).Description = DescriptionText.Text
     
        '*-------------------------------------*
        '* Check if node is action an not root *
        '*-------------------------------------*
        
        If OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).IsAction And (TreeView1.SelectedItem.Index <> 1) Then
    
            '**** Fill OA values
    
            OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).Perception = PerceptionType.ListIndex
            OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).ChanceOfSuccess = ChanceOfHappening.ListIndex
            OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).Headline = HeadlineText.Text
            OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).PowerPointSlidePath = PowerPointSlidePath.Text
        
        Else
            
            If TreeView1.SelectedItem.Index = 1 Then
        
                '**** Special case for root node
        
                OptionsAndActions.Item(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key).PowerPointSlidePath = PowerPointSlidePath.Text
        
            End If
        
        End If
    
    End If

End Sub
