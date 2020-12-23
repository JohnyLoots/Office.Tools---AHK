#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#MaxThreadsPerHotkey 2

!r::Reload
!q::ExitApp


;                           Mail To
;---------------------------------------------------------------------
!^m::
InputBox, Mail_to_var,Email Address., Enter Email Address? 
if (Mail_to_var = ""){
	Soundplay, C:\Windows\Media\Speech Off.wav
}
else
	InputBox, Subject_var,Subject, Please Enter Subject for Email.
	Soundplay, C:\Windows\Media\Speech On.wav
	Run, mailto:%Mail_to_var%?subject=%Subject_var%
return
;---------------------------------------------------------------------



;                         Google Search
;---------------------------------------------------------------------
!^g::
InputBox, Google_search_var,Google Search., Enter what you would like to search ? 
if (Google_search_var = ""){
	Soundplay, C:\Windows\Media\Speech Off.wav
}
else
	Soundplay, C:\Windows\Media\Speech On.wav
	Run, http://www.google.com/search?q=%Google_search_var%
return
;---------------------------------------------------------------------



;                     Open Notepad And Save Desktop
;---------------------------------------------------------------------
!^n:: 
Run Notepad.exe %A_Desktop%\New text file.txt
return 
;---------------------------------------------------------------------




;                     Open Sublime Text 3 
;---------------------------------------------------------------------
!^s::
sublime_path = C:\Program Files\Sublime Text 3\sublime_text.exe
Run, %sublime_path%
return
;---------------------------------------------------------------------

;                     Open WireShark
;---------------------------------------------------------------------
!^w::
wireshark_path = C:\Program Files\Wireshark\Wireshark.exe
Run, %wireshark_path%
return
;---------------------------------------------------------------------



;                     Create New Excell Sheet
;---------------------------------------------------------------------
^o::
FileSelectFile, SelectedFileOutput, 8, C:\, Select a workbook.,All Excel Files(*.xl; *.xlsx; *.xlsm; *.xlsb; *.xlam; *.xltx; *.xls; *.xlt; *.htm; *.html; *.mht; *.mhtml; *.xml; *.xla; *.xlm; *.xlw; *.odc; *.ods)
 
	Bookname := SelectedFileOutput


oExcell := ComObjCreate("Excel.Application")
oExcell.Visible := True
oExcell.Workbooks.Add
oExcell.ActiveWorkbook.SaveAs(Bookname)
oExcell_Workbook := oExcell.Workbooks.Open(Bookname)

InputBox,Nameoutputvar, SheetTitle., Enter Sheet Title ? 
if errorlevel 
	Name_var := "Insert Title Here."
else if (Nameoutputvar = "")
	Name_var := Nameoutputvar 
else
	Name_var := Nameoutputvar

oExcell.range("B2").value := Name_var 
oExcell.range("A4").value := "Datum"
oExcell.range("B4").value := "Staatnr"
oExcell.range("C4").value := "Beskryfwing"
oExcell.range("D4").value := "Beskryfwing"
oExcell.range("E4").value := "Debiet"
oExcell.range("F4").value := "Krediet"
oExcell.range("G4").value := "Saldo"
oExcell.range("G6").value := "=G5+F6-E6"
oExcell.range("G7").value := "=G6+F7-E7"
oExcell.range("G8").value := "=G7+F8-E8"
oExcell.range("G9").value := "=G8+F9-E9"
oExcell.range("G10").value := "=G9+F10-E10"
oExcell.range("G11").value := "=G10+F11-E11"
oExcell.range("G12").value := "=G11+F12-E12"
oExcell.range("G13").value := "=G12+F13-E13"
oExcell.range("G14").value := "=G13+F14-E14"
oExcell.range("G15").value := "=G14+F15-E15"

oExcell.range("E17").value := "=SUM(E5:E16)" 
oExcell.range("F17").value := "=SUM(F5:F16)"
oExcell.range("G17").value := "=SUM(G5:G16)" 


oExcell.range("B2").font.italic := true
oExcell.range("B2").font.size := 20
oExcell.range("A4").font.bold := true
oExcell.range("A4").font.size := 12
oExcell.range("B4").font.bold := true
oExcell.range("B4").font.size := 12
oExcell.range("C4").font.bold := true
oExcell.range("C4").font.size := 12
oExcell.range("D4").font.bold := true
oExcell.range("D4").font.size := 12
oExcell.range("E4").font.bold := true
oExcell.range("E4").font.size := 12
oExcell.range("F4").font.bold := true
oExcell.range("F4").font.size := 12
oExcell.range("G4").font.bold := true
oExcell.range("G4").font.size := 12

Soundplay, C:\Windows\Media\Tada.wav
oExcell.ActiveWorkbook.Save

return
;---------------------------------------------------------------------
