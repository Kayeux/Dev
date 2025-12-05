Run("notepad.exe")
WinWaitActive("Untitled - Notepad")
Send("Test Thing")
WinClose("Untitled - Notepad")