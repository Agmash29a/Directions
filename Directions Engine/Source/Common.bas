Attribute VB_Name = "Common"
Public IsAction() As Boolean
Public Description() As String
Public Headline() As String
Public ChanceOfSuccess() As Integer
Public Perception() As Integer
Public Parent() As Integer
Public NodeTitles() As String
Public PowerPointSlidePaths() As String
Public NumberOfNodes As Long

Public ScenarioLoaded As Boolean
Public StartGame As Boolean

Public CurrentNode As Integer

Public Const HOVER_COLOR_RED = 45
Public Const HOVER_COLOR_GREEN = 153
Public Const HOVER_COLOR_BLUE = 175

Public Const FORM_COLOR_RED = 255
Public Const FORM_COLOR_GREEN = 255
Public Const FORM_COLOR_BLUE = 233

Public Const KEY_SPACE = 32
Public Const KEY_ENTER = 13
Public Const KEY_P = 112
Public Const KEY_1 = 49
Public Const KEY_2 = 50
Public Const KEY_3 = 51
Public Const KEY_4 = 52
Public Const KEY_5 = 53
Public Const KEY_6 = 54

Public OptionPage As Boolean
