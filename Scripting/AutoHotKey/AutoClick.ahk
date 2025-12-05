#Requires AutoHotkey v2.0
CapsLock:: {  ; CapsLock = Auto-click
 Static on := False
 If on := !on
      SetTimer(Click, 50), Click()
 Else SetTimer(Click, 0)
}