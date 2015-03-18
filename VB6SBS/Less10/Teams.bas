Attribute VB_Name = "Module1"
Sub AddName(Team$, ReturnString$)
    Prompt$ = "Enter a " & Team$ & " employee."
    Nm$ = InputBox(Prompt$, "Input Box")
    WrapCharacter$ = Chr(13) + Chr(10)
    ReturnString$ = Nm$ & WrapCharacter$
End Sub


