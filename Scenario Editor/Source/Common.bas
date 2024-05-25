Attribute VB_Name = "Common"
Public Errors As Boolean
Public RootNodeCount As Integer
Public KeyCounter As Integer
Public OptionsAndActions As New OptionActions
Public ScenarioTitleVar As String
Public FirstQuestionVar As String

Public Const MAX_OPTION_NODES = 6

Public Const PERCEPTION_TYPE_BAD = 0
Public Const PERCEPTION_TYPE_GOOD = 1

Public Const CHANCE_VERY_HIGH = 0
Public Const CHANCE_HIGH = 1
Public Const CHANCE_AVERAGE = 2
Public Const CHANCE_LOW = 3
Public Const CHANCE_VERY_LOW = 4

Public Const PICTURE_ACTION_NODE = 1
Public Const PICTURE_OPTION_NODE = 2
Public Const PICTURE_GAME_NODE = 3

Public Const WIDTH_ORIGINAL_FORM = 10935
Public Const WIDTH_ORIGINAL_FRAME1 = 4335
Public Const WIDTH_ORIGINAL_DESCRIPTION_TEXT = 4095

Public Const LEFT_ORIGINAL_FRAME3 = 6120

Public Const HEIGHT_MINIMUM_RESIZE = 3000
Public Const HEIGHT_ORIGINAL_TREEVIEW = 4935
Public Const HEIGHT_ORIGINAL_FRAMES = 5040
Public Const HEIGHT_NON_TREEVIEW = 1995

Sub TreeViewBuild(TV As TreeView, _
    IsAction() As Boolean, Description() As String, Headline() As String, ChanceOfSuccess() As Integer, _
    Perception() As Integer, Parent() As Integer, NodeTitles() As String, PowerPointSlidePaths() As String, Optional ByVal StartNode As Integer)
   
    '*----------------*
    '*                *
    '* Build Treeview *
    '* ~~~~~~~~~~~~~~ *
    '*                *
    '*----------------*
          
    Dim nd As Integer
    Dim nde As Node
    Dim oa As OptionAction
    Dim keytext As String

    '*----------------------------*
    '* Check if node is root node *
    '*----------------------------*

    If StartNode = 0 Then
        ' if StartNode is omitted, start from the first root node
       nd = 1
       
    Else
    
       nd = StartNode
    
    End If
   
   '*----------------*
   '* Build treeview *
   '*----------------*
   
    keytext = "key" & nd
    
    Set oa = OptionsAndActions.Add(IsAction(nd), Description(nd), Headline(nd), ChanceOfSuccess(nd), Perception(nd), PowerPointSlidePaths(nd), keytext)
         
    If nd = 1 Then
       
        '**** Root node
       
        Set nde = TV.Nodes.Add(, , keytext, NodeTitles(nd), PICTURE_GAME_NODE)
    
    Else
    
        '**** Other Node
        
        '*-----------------------------------*
        '* Check if node is Action or option *
        '*-----------------------------------*
   
        If IsAction(nd) Then
        
            '**** Action Node
        
            Set nde = TV.Nodes.Add("key" & Parent(nd), tvwChild, keytext, NodeTitles(nd), PICTURE_ACTION_NODE)
        
        Else
        
            '**** Option Node
        
            Set nde = TV.Nodes.Add("key" & Parent(nd), tvwChild, keytext, NodeTitles(nd), PICTURE_OPTION_NODE)
        
        End If
        
    End If
   
    '*-------------------------------*
    '* Find all child nodes and call *
    '*-------------------------------*
    
    For i = 1 To UBound(Parent)
    
        If (Parent(i) = nd) And (i <> nd) Then
        
            Call TreeViewBuild(TV, IsAction, Description, Headline, ChanceOfSuccess, Perception, Parent, NodeTitles, PowerPointSlidePaths, i)
        
        End If
          
    Next
  
    KeyCounter = KeyCounter + 1
    
    '**** Clean up
    
    Set nde = Nothing
    Set oa = Nothing

End Sub

Sub TreeViewParse(TV As TreeView, oa As OptionActions, Optional StartNode As Node, Optional OnlyVisible As Boolean)
   
    '*-------------------*
    '*                   *
    '* Parse Collections *
    '* ~~~~~~~~~~~~~~~~~ *
    '*                   *
    '*-------------------*
    
    Dim nd As Node, childND As Node
 
    ' exit if there are no nodes
    If TV.Nodes.Count = 0 Then Exit Sub
    If StartNode Is Nothing Then
        ' if StartNode is omitted, start from the first root node
        Set nd = TV.Nodes(1).Root.FirstSibling
    Else
        Set nd = StartNode
    End If
    ' output the starting node
    
    '*-----------------------------------------------------*
    '* Check all children for Actions if this is an Option *
    '*-----------------------------------------------------*
    
    If oa.Item(nd.Key).IsAction = False Then
    
        '**** it's an option, check if it has at least 1 child
        
        If nd.Children = 0 Then
        
            Errors = True
        
        End If
        
    End If
    
    '****
 
    ' then call recursively this routine to output all child nodes
    ' if OnlyVisible=Tree, do this only if this node is expanded
    If nd.Children And (nd.Expanded Or OnlyVisible = False) Then
       
        Set childND = nd.Child
        For i = 1 To nd.Children
            Call TreeViewParse(TV, oa, childND, OnlyVisible)
            Set childND = childND.Next
        Next
        
    End If
  
End Sub


Function TreeViewToString(TV As TreeView, Optional StartNode As Node, Optional OnlyVisible As Boolean) As String
    Dim nd As Node, childND As Node
    Dim res As String, i As Long
    Static Level As Integer
    
    ' exit if there are no nodes
    If TV.Nodes.Count = 0 Then Exit Function
    If StartNode Is Nothing Then
        ' if StartNode is omitted, start from the first root node
        Set nd = TV.Nodes(1).Root.FirstSibling
    Else
        Set nd = StartNode
    End If
    ' output the starting node
    res = String$(Level, vbTab) & nd.Text & vbCrLf
    
    ' then call recursively this routine to output all child nodes
    ' if OnlyVisible=Tree, do this only if this node is expanded
    If nd.Children And (nd.Expanded Or OnlyVisible = False) Then
        Level = Level + 1
        Set childND = nd.Child
        For i = 1 To nd.Children
            res = res & TreeViewToString(TV, childND, OnlyVisible)
            Set childND = childND.Next
        Next
        Level = Level - 1
    End If
    
    ' if we are parsing the whole tree, we must account for multiple roots
    If StartNode Is Nothing Then
        Set nd = nd.Next
        Do Until nd Is Nothing
            res = res & TreeViewToString(TV, nd, OnlyVisible)
            Set nd = nd.Next
        Loop
    End If
    
    TreeViewToString = res
End Function

Public Function GetNewKey() As String

    KeyCounter = KeyCounter + 1

    GetNewKey = Str(KeyCounter)

End Function
