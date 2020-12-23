#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

$!r:: 
 OldClipboard := ClipboardAll
  Clipboard = ;clears the Clipboard
  SendInput {Left}+{Right 2}
  SendInput, ^c
  ClipWait 0 ;pause for Clipboard data
  If ErrorLevel
  {
    MsgBox, No text selected!
  }
  SwappedLetters := SubStr(Clipboard,2) . SubStr(Clipboard,1,1)
  SendInput, %SwappedLetters%
  SendInput {Left}
 Clipboard := OldClipboard
Return
