VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarIsAction As Boolean 'local copy
Private mvarDescription As String 'local copy
Private mvarHeadline As String 'local copy
Private mvarChanceOfSuccess As Integer 'local copy
Private mvarPerception As Integer 'local copy
Private mvarPowerPointSlidePath As String

Public Property Let PowerPointSlidePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PowerPointSlidePath = "5"
    mvarPowerPointSlidePath = vData
End Property

Public Property Get PowerPointSlidePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PowerPointSlidePath
    PowerPointSlidePath = mvarPowerPointSlidePath
End Property

Public Property Let Perception(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Perception = 5
    mvarPerception = vData
End Property


Public Property Get Perception() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Perception
    Perception = mvarPerception
End Property



Public Property Let ChanceOfSuccess(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChanceOfSuccess = 5
    mvarChanceOfSuccess = vData
End Property


Public Property Get ChanceOfSuccess() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChanceOfSuccess
    ChanceOfSuccess = mvarChanceOfSuccess
End Property



Public Property Let Headline(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Headline = 5
    mvarHeadline = vData
End Property


Public Property Get Headline() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Headline
    Headline = mvarHeadline
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let IsAction(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsAction = 5
    mvarIsAction = vData
End Property


Public Property Get IsAction() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsAction
    IsAction = mvarIsAction
End Property



