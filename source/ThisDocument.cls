VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_Open()
    CustomizationContext = ActiveDocument
    KeyBindings.Add wdKeyCategoryMacro, "WordleGuess", BuildKeyCode(wdKeyReturn)
    
    InitWordLists
End Sub
