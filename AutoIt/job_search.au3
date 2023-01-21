#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Script Function:
	Searchs for jobs automatically

#ce ----------------------------------------------------------------------------
; INCLUDES
; Includes all IE constants
#Include "IE.AU3"
#include <MsgBoxConstants.au3>

; AUTO IT OPTIONS SETTINGS
; This tells AutoIt how to match window titles.
AutoItSetOption("WinTitleMatchMode", 2)
; See the help file for details of other matching modes.

; This line just complicates my explanation of what is going on.
AutoItSetOption("WinDetectHiddenText", 1)
; Wait until the end of the script for more info on this.

; This tells AutoIt how long to pause
; after a successful window-related operation.
AutoItSetOption("WinWaitDelay", 1); (milliseconds)
; See the help file for more details.

; This helps during the debug of the script
AutoItSetOption("TrayIconDebug", 1);0-off

; Get the mouse out of the way.
MouseMove(0, 0, 10)
; If the mouse happens to be where a link could appear in the
; web page, then it could mess up this script.

; VARIABLES
; Declare object variable related to IE, starts IE and search for webpage login webpage.
Local $oIE = _IECreate("https://www2.webpage.cl/login")

; Set multiple variables that resumes the biggest part of the different URL's considering the keyword search
Local $sAsistente = "URL_WEBPAGE"
Local $sAdministrativo = "URL_WEBPAGE"
Local $sComercial = "URL_WEBPAGE"
Local $sEjecutivo = "URL_WEBPAGE"
Local $sSecretario = "URL_WEBPAGE"
Local $sAuxiliar = "URL_WEBPAGE"
Local $sAnalista = "URL_WEBPAGE"
Local $sRecepcionista = "URL_WEBPAGE"
Local $sCoordinador = "URL_WEBPAGE"

;Hot key to Stop the script
HotKeySet( "{END}", "Stop" )

;Function definition to stop de script
Func Stop ()
	Exit
EndFunc

; Script Engine
 webpage ()

 Func webpage ()

; wait for a window with "Acceso a mi cuenta" in the title
 WinWait("Acceso a mi cuenta", "")

; Active that window - just in case
 WinActivate("Acceso a mi cuenta", "")

; Wait until that window is active
 WinWaitActive("Acceso a mi cuenta", "")

 Sleep(5000)

; TAB Search for the fields, fill it and log in
 Send("{TAB 16}email@gmail.com{TAB}password{TAB 2}{ENTER}")

 Sleep(5000)

; Search for the first keyword
_IENavigate( $oIE, $sAdministrativo & 1)

; Waiting and activate fuctions used before
WinWait("administrativo", "")
WinActivate("administrativo", "")
WinWaitActive("administrativo", "")

; TAB Search for the first link
Do; This loop repeats the code between the Do and Until lines.
    Send("{TAB}")
    Sleep(500); <<Slows the loop down, change the speed here.
Until WinExists("administrativo", "administrativo")

 EndFunc