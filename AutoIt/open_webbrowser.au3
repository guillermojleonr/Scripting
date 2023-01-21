#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Guillermo Leon

 Script Function:
	Searchs for jobs in webpage.cl

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

webpage ()
Func webpage ()
	;Run Chrome
	Local $iOPE = Run("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", "", @SW_SHOWMAXIMIZED)

	;Wait until the windows is active
	WinWaitActive ("Google Chrome")
	Sleep(3000)

	Send("hola")

EndFunc