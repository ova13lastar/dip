; #INDEX# =======================================================================================================================
; Title .........: DIP
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
#AutoIt3Wrapper_Res_ProductName=DIP
#AutoIt3Wrapper_Res_Description=Dématérialisation des Impressions PROGRES
#AutoIt3Wrapper_Res_ProductVersion=0.0.6
#AutoIt3Wrapper_Res_FileVersion=0.0.6
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/APPLINAT
#AutoIt3Wrapper_Res_LegalCopyright=yann.daniel@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Icon="static\icon.ico"
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Run_Au3Stripper=N
#Au3Stripper_Parameters=/MO /RSLN
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes YD
#include "D:\Autoit_dev\Include\YDGVars.au3"
#include "D:\Autoit_dev\Include\YDLogger.au3"
#include "D:\Autoit_dev\Include\YDTool.au3"
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
; Includes
#include <String.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
OnAutoItExitRegister("_RestoreOnError")
OnAutoItExitRegister("_YDTool_ExitApp")
; ===============================================================================================================================

; #VARIABLES GLOBALES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("FileVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)

_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On gere l'affichage de l'icone dans le tray
TraySetIcon(_YDGVars_Get("sAppIconPath"))
TraySetToolTip(_YDGVars_Get("sAppTitle"))
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
;------------------------------
; On initialisation d'autres variables globales (dont celles du conf.ini)
Global $g_sSiteNetworkPath		= "D:\"
Global $g_iLastLineTechLogFile  = 1
Global $g_iLastLineNticLogFile  = 1
Global $g_iLastLineInjLogFile  = 1
Global $g_IniSectionName = ""
; --- INI : printers
$g_IniSectionName = "printers"
Global $g_sPdfCreatorPrinter = _YDTool_GetAppConfValue($g_IniSectionName, "printer_name")
Global $g_sRegexLocalPrinter = _YDTool_GetAppConfValue($g_IniSectionName, "regex_local_printer")
; --- INI : progres
$g_IniSectionName = "progres"
Global $g_sProgresMainPath = _YDTool_GetAppConfValue($g_IniSectionName, "main_path")
Global $g_sProgresExeFileName = _YDTool_GetAppConfValue($g_IniSectionName, "exe_filename")
Global $g_sProgresExeFilePath = $g_sProgresMainPath & "\" & $g_sProgresExeFileName
Global $g_sProgresLogPath = _YDTool_GetAppConfValue($g_IniSectionName, "log_path")
Global $g_sProgresNsReportPath = _YDTool_GetAppConfValue($g_IniSectionName, "nsreport_path")
Global $g_sProgresTechLogFilePath = $g_sProgresLogPath & "\TECH_" & @YEAR & @MON & @MDAY & ".LOG"
Global $g_sProgresNticLogFilePath = $g_sProgresLogPath & "\NTIC_" & @YEAR & @MON & @MDAY & ".LOG"
Global $g_sProgresInjLogFilePath = $g_sProgresLogPath & "\INJ_" & @YEAR & @MON & @MDAY & ".LOG"
_YDLogger_Var("$g_sProgresExeFilePath", $g_sProgresExeFilePath)
_YDLogger_Var("$g_sProgresTechLogFilePath", $g_sProgresTechLogFilePath)
_YDLogger_Var("$g_sProgresNticLogFilePath", $g_sProgresNticLogFilePath)
_YDLogger_Var("$g_sProgresInjLogFilePath", $g_sProgresInjLogFilePath)
; --- INI : progres_ouverture
$g_IniSectionName = "progres_ouverture"
Global $g_sProgresOuvertureArrasPath = _YDTool_GetAppConfValue($g_IniSectionName, "arras_path")
Global $g_sProgresOuvertureLensPath = _YDTool_GetAppConfValue($g_IniSectionName, "lens_path")
Global $g_sProgresOuvertureAutoOpenOutPutDir = _YDTool_GetAppConfValue($g_IniSectionName, "auto_open_output_dir")
Global $g_sProgresOuvertureWindowTitle = _YDTool_GetAppConfValue($g_IniSectionName, "window_title")
; --- INI : progres_liasses
$g_IniSectionName = "progres_liasses"
Global $g_sProgresLiassesArrasPath = _YDTool_GetAppConfValue($g_IniSectionName, "arras_path")
Global $g_sProgresLiassesLensPath = _YDTool_GetAppConfValue($g_IniSectionName, "lens_path")
Global $g_sProgresLiassesAutoOpenOutPutFile = _YDTool_GetAppConfValue($g_IniSectionName, "auto_open_output_file")
; MCO
Global $g_sProgresLiassesMcoDatFileName = _YDTool_GetAppConfValue($g_IniSectionName, "dat_filename")
Global $g_sProgresLiassesMcoDatFilePath = $g_sProgresNsReportPath & "\" & $g_sProgresLiassesMcoDatFileName
Global $g_sProgresLiassesMcoDatFileDateTime = (FileExists($g_sProgresLiassesMcoDatFilePath)) ? _ArrayToString(FileGetTime($g_sProgresLiassesMcoDatFilePath)) : 0
Global $g_bProgresLiassesMcoDatFileChanged = False
_YDLogger_Var("$g_sProgresLiassesMcoDatFilePath", $g_sProgresLiassesMcoDatFilePath)
_YDLogger_Var("$g_sProgresLiassesMcoDatFileDateTime", $g_sProgresLiassesMcoDatFileDateTime)
; --- INI : progres_injecteurs
$g_IniSectionName = "progres_injecteurs"
Global $g_sProgresInjecteursEtatArrasPath = _YDTool_GetAppConfValue($g_IniSectionName, "arras_path")
Global $g_sProgresInjecteursEtatLensPath = _YDTool_GetAppConfValue($g_IniSectionName, "lens_path")
Global $g_sProgresInjecteursAutoOpenOutPutFile = _YDTool_GetAppConfValue($g_IniSectionName, "auto_open_output_file")
; Relance
Global $g_sProgresInjecteursEtatRelanceDatFileName = _YDTool_GetAppConfValue($g_IniSectionName, "etat_relance_dat_filename")
Global $g_sProgresInjecteursEtatRelanceDatFilePath = $g_sProgresNsReportPath & "\" & $g_sProgresInjecteursEtatRelanceDatFileName
Global $g_sProgresInjecteursEtatRelanceDatFileDateTime = (FileExists($g_sProgresInjecteursEtatRelanceDatFilePath)) ? _ArrayToString(FileGetTime($g_sProgresInjecteursEtatRelanceDatFilePath)) : 0
Global $g_bProgresInjecteursEtatRelanceDatFileChanged = False
_YDLogger_Var("$g_sProgresInjecteursEtatRelanceDatFilePath", $g_sProgresInjecteursEtatRelanceDatFilePath)
_YDLogger_Var("$g_sProgresInjecteursEtatRelanceDatFileDateTime", $g_sProgresInjecteursEtatRelanceDatFileDateTime)
; Rejet
Global $g_sProgresInjecteursEtatRejetDatFileName = _YDTool_GetAppConfValue($g_IniSectionName, "etat_rejet_dat_filename")
Global $g_sProgresInjecteursEtatRejetDatFilePath = $g_sProgresNsReportPath & "\" & $g_sProgresInjecteursEtatRejetDatFileName
Global $g_sProgresInjecteursEtatRejetDatFileDateTime = (FileExists($g_sProgresInjecteursEtatRejetDatFilePath)) ? _ArrayToString(FileGetTime($g_sProgresInjecteursEtatRejetDatFilePath)) : 0
Global $g_bProgresInjecteursEtatRejetDatFileChanged = False
_YDLogger_Var("$g_sProgresInjecteursEtatRejetDatFilePath", $g_sProgresInjecteursEtatRejetDatFilePath)
_YDLogger_Var("$g_sProgresInjecteursEtatRejetDatFileDateTime", $g_sProgresInjecteursEtatRejetDatFileDateTime)
; OK
Global $g_sProgresInjecteursEtatOkDatFileName = _YDTool_GetAppConfValue($g_IniSectionName, "etat_ok_dat_filename")
Global $g_sProgresInjecteursEtatOkDatFilePath = $g_sProgresNsReportPath & "\" & $g_sProgresInjecteursEtatOkDatFileName
Global $g_sProgresInjecteursEtatOkDatFileDateTime = (FileExists($g_sProgresInjecteursEtatOkDatFilePath)) ? _ArrayToString(FileGetTime($g_sProgresInjecteursEtatOkDatFilePath)) : 0
Global $g_bProgresInjecteursEtatOkDatFileChanged = False
_YDLogger_Var("$g_sProgresInjecteursEtatOkDatFilePath", $g_sProgresInjecteursEtatOkDatFilePath)
_YDLogger_Var("$g_sProgresInjecteursEtatOkDatFileDateTime", $g_sProgresInjecteursEtatOkDatFileDateTime)
; AV
Global $g_sProgresInjecteursEtatAvDatFileName = _YDTool_GetAppConfValue($g_IniSectionName, "etat_av_dat_filename")
Global $g_sProgresInjecteursEtatAvDatFilePath = $g_sProgresNsReportPath & "\" & $g_sProgresInjecteursEtatAvDatFileName
Global $g_sProgresInjecteursEtatAvDatFileDateTime = (FileExists($g_sProgresInjecteursEtatAvDatFilePath)) ? _ArrayToString(FileGetTime($g_sProgresInjecteursEtatAvDatFilePath)) : 0
Global $g_bProgresInjecteursEtatAvDatFileChanged = False
_YDLogger_Var("$g_sProgresInjecteursEtatAvDatFilePath", $g_sProgresInjecteursEtatAvDatFilePath)
_YDLogger_Var("$g_sProgresInjecteursEtatAvDatFileDateTime", $g_sProgresInjecteursEtatAvDatFileDateTime)
; ---
Global $g_sDefaultPrinter = _YDTool_GetDefaultPrinter(@ComputerName)
Global $g_sDefaultPrinterName = StringRegExpReplace($g_sDefaultPrinter, $g_sRegexLocalPrinter, "")
Global $g_sSite = _YDTool_GetHostSite(@ComputerName)
_YDLogger_Var("$g_sDefaultPrinterName", $g_sDefaultPrinterName)
;------------------------------
; On reinstalle systematiquement l'imprimante pdfCreator si utilisateur connecte
Global $g_sLoggerUserName = _YDTool_GetHostLoggedUserName(@ComputerName)
If $g_sLoggerUserName <> "" Then
	_InstallPdfCreatorPrinter()
EndIf
; On verifie que l'installation s'est bien passee
If _IsPdfCreatorPrinterInstalled() Then
	_YDLogger_Log("Imprimante " & $g_sPdfCreatorPrinter & " installee pour utilisateur : " & $g_sLoggerUserName)
Else
	_YDLogger_Error("Imprimante " & $g_sPdfCreatorPrinter & " non installee malgre la tentative d'installation !")
EndIf
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
	Global $iMsg = TrayGetMsg()
	Select
		Case $iMsg = $idTrayExit
			_YDTool_ExitConfirm()
		Case $iMsg = $idTrayAbout
			_YDTool_GUIShowAbout()
		Case Else
			_ModuleInjecteurs()
			_ModuleLiasses()
			_ModuleOuvertures()
	EndSelect
	;------------------------------
	Sleep(10)
WEnd
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier le contexte
; Syntax ........: _CheckContext()
; Parameters ....:
; Return values .: True si OK / False si KO
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 09/12/2019
; Notes .........:
;================================================================================================================================
Func _CheckContext()
	Local $sFuncName = "_CheckContext"
	local $bReturn = True
	; On verifie qu'un utilisateur est connecte
	If $g_sLoggerUserName = "" Then
		_YDLogger_Log("Aucun utilisateur connecte !", $sFuncName, 1)
		$bReturn = False
	EndIf
	; On verifie que l'imprimante pdfCreator est bien installee
	If Not _IsPdfCreatorPrinterInstalled() Then
		_YDLogger_Log("Imprimante " & $g_sPdfCreatorPrinter & " non installee !", $sFuncName, 1)
		$bReturn = False
	EndIf
	_YDLogger_Var("$bReturn", $bReturn, $sFuncName, 2)
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement des impressions de liasses
; Syntax ........: _ModuleLiasses()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _ModuleLiasses()
	Local $sFuncName = "_ModuleLiasses"
	Local $i
	; On verifie le contexte
	_CheckContext()
	; On ne travaille que si PROGRES est lance
	If ProcessExists($g_sProgresExeFileName) Then
		; On verifie si l'impression a démarré
		If _IsPrintStartFromMcoDatFile() Then
			_YDLogger_Log("Impression liasse démarrée !!!!", $sFuncName)
			; On suspend PROGRES
			_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, True)
			Sleep(1000)
			; On tente de lire le fichier GCO_MCO.DAT
			Local $aMcoDat
			_YDLogger_Log("Contenu du fichier " & $g_sProgresLiassesMcoDatFilePath & " :", $sFuncName, 1)
			If _FileReadToArray($g_sProgresLiassesMcoDatFilePath, $aMcoDat) = 0 Then
				_YDLogger_Error("Impossible de lire le fichier", $sFuncName)
				_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, False)
				Return False
			EndIf
			; Si OK, on recupere les donnees
			For $i = 0 To $aMcoDat[0]
				Local $aMcoDatVar = StringSplit($aMcoDat[$i], "=")
				_YDLogger_Log($aMcoDat[$i], $sFuncName, 1)
				Switch $aMcoDatVar[1]
					Case "LIASSE"
						Local $sLiasse = $aMcoDatVar[2]
					Case "CAISSE"
						Local $sCaisse = $aMcoDatVar[2]
					Case "CENTRE"
						Local $sUGE = $aMcoDatVar[2]
					Case "AGENT"
						Local $sAgent = $aMcoDatVar[2]
					Case "ACTION"
						Local $bEcheancierAuto = False
						If $aMcoDatVar[2] == "Trait. éch auto" Then
							$bEcheancierAuto = True
						EndIf
				EndSwitch
			Next
			_YDLogger_Var("$bEcheancierAuto", $bEcheancierAuto)
			If $bEcheancierAuto Then
				_YDLogger_Log("Echeancier auto : pas de bascule", $sFuncName, 1)
				; On reactive PROGRES
				_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, False)
			Else
				_YDLogger_Log("Traitement classique : on doit basculer ...", $sFuncName, 1)
				Local $sMcoDatFileDateTime = FileGetTime($g_sProgresLiassesMcoDatFilePath)
				Local $sDate = $sMcoDatFileDateTime[0] & $sMcoDatFileDateTime[1] & $sMcoDatFileDateTime[2]
				Local $sTime = $sMcoDatFileDateTime[3] & $sMcoDatFileDateTime[4] & $sMcoDatFileDateTime[5]
				$g_sSiteNetworkPath = ($g_sSite = "ARRAS") ? $g_sProgresLiassesArrasPath : $g_sProgresLiassesLensPath
				_YDLogger_Var("$g_sSiteNetworkPath", $g_sSiteNetworkPath)
				Local $sAutosaveDirectory = $g_sSiteNetworkPath & "\" & $sDate & "\"
				Local $sAutosaveFilename = $sDate & "-" & $sTime & "_" & $sCaisse & "_" & $sUGE & "_" & $sLiasse & "_" & $sAgent
				; On modifie le registre pour modifier le Path et le nom du fichier
				_YDLogger_Log("Modification du registre", $sFuncName, 1)
				_YDLogger_Var("$sAutosaveDirectory", $sAutosaveDirectory, $sFuncName, 1)
				_YDLogger_Var("$sAutosaveFilename", $sAutosaveFilename, $sFuncName, 1)
				RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveDirectory", "REG_SZ", $sAutosaveDirectory)
				RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveFilename", "REG_SZ", $sAutosaveFilename)
				RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveStartStandardProgram", "REG_SZ", $g_sProgresLiassesAutoOpenOutPutFile)
				; On bascule sur PDFCreator
				$i = 0
				_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
				While _YDTool_GetDefaultPrinter(@ComputerName) <> $g_sPdfCreatorPrinter
					$i += 1
					Sleep(100)
					_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
					If $i > 20 Then
						_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule impossible vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
						Return False
					EndIf
				WEnd
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
				; On reactive PROGRES
				_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, False)
				; L'impression se lance ...
				Sleep(1000)
				; On verifie si l'impression est terminee
				If _IsPrintStopFromTechLogFile() Then
					_YDLogger_Log("Impression liasse terminée !", $sFuncName)
					; On retourne sur l'imprimante par defaut
					_YDTool_SetDefaultPrinter($g_sDefaultPrinter)
					_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Retour sur imprimante : " & $g_sDefaultPrinterName, 5000)
				EndIf
			EndIf
		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement des impressions pour les injecteurs
; Syntax ........: _ModuleInjecteurs()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _ModuleInjecteurs()
	Local $sFuncName = "_ModuleInjecteurs"
	;Local $i
	Local $sTestsPath
	; On verifie le contexte
	_CheckContext()
	; On ne travaille que si PROGRES est lance
	If ProcessExists($g_sProgresExeFileName) Then
		; ---------------------------
		; On verifie si l'impression d'un état a démarré
		If _IsPrintStartFromEtatRelanceDatFile() Then
			_YDLogger_Log("Impression état RELANCE démarrée !!!!", $sFuncName)
			; On patiente tant que le fichier n'est pas libéré ou 5 secondes
			For $i = 1 To 500
				If FileOpen($g_sProgresInjecteursEtatRelanceDatFilePath) <> -1 Then ExitLoop
				Sleep(10)
			Next
			; On copie le fichier
			$sTestsPath = "C:\APPLILOC\DIP\tests\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & "_RELANCE.DAT"
			_YDTool_CopyFile($g_sProgresInjecteursEtatRelanceDatFilePath, $sTestsPath)
		EndIf
		If _IsPrintStartFromEtatRejetDatFile() Then
			_YDLogger_Log("Impression état REJET démarrée !!!!", $sFuncName)
			; On patiente tant que le fichier n'est pas libéré ou 5 secondes
			For $i = 1 To 500
				If FileOpen($g_sProgresInjecteursEtatRejetDatFilePath) <> -1 Then ExitLoop
				Sleep(10)
			Next
			; On copie le fichier
			$sTestsPath = "C:\APPLILOC\DIP\tests\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & "_REJET.DAT"
			_YDTool_CopyFile($g_sProgresInjecteursEtatRejetDatFilePath, $sTestsPath)
		EndIf
		If _IsPrintStartFromEtatOkDatFile() Then
			_YDLogger_Log("Impression état OK démarrée !!!!", $sFuncName)
			; On patiente tant que le fichier n'est pas libéré ou 5 secondes
			For $i = 1 To 500
				If FileOpen($g_sProgresInjecteursEtatOkDatFilePath) <> -1 Then ExitLoop
				Sleep(10)
			Next
			; On copie le fichier
			$sTestsPath = "C:\APPLILOC\DIP\tests\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & "_OK.DAT"
			_YDTool_CopyFile($g_sProgresInjecteursEtatOkDatFilePath, $sTestsPath)
		EndIf
		If _IsPrintStartFromEtatAvDatFile() Then
			_YDLogger_Log("Impression état AV démarrée !!!!", $sFuncName)
			; On patiente tant que le fichier n'est pas libéré ou 5 secondes
			For $i = 1 To 500
				If FileOpen($g_sProgresInjecteursEtatAvDatFilePath) <> -1 Then ExitLoop
				Sleep(10)
			Next
			; On copie le fichier
			$sTestsPath = "C:\APPLILOC\DIP\tests\" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "-" & @MSEC & "_AV.DAT"
			_YDTool_CopyFile($g_sProgresInjecteursEtatAvDatFilePath, $sTestsPath)
		EndIf
		; ----------------------------
		If ($g_bProgresInjecteursEtatRelanceDatFileChanged Or $g_bProgresInjecteursEtatRejetDatFileChanged Or $g_bProgresInjecteursEtatOkDatFileChanged Or $g_bProgresInjecteursEtatAvDatFileChanged) Then
			; On suspend PROGRES
			_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, True)
			Sleep(1000)
			; On détermine le site et le nom du fichier
			$g_sSiteNetworkPath = ($g_sSite = "ARRAS") ? $g_sProgresLiassesArrasPath : $g_sProgresLiassesLensPath
			_YDLogger_Var("$g_sSiteNetworkPath", $g_sSiteNetworkPath)
			Local $sAutosaveDirectory = $g_sSiteNetworkPath & "\" & @YEAR & @MON & @MDAY & "\"
			Local $sAutosaveFilename = @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & "_etat"
			; On modifie le registre pour modifier le Path et le nom du fichier
			_YDLogger_Log("Modification du registre", $sFuncName, 1)
			_YDLogger_Var("$sAutosaveDirectory", $sAutosaveDirectory, $sFuncName, 1)
			_YDLogger_Var("$sAutosaveFilename", $sAutosaveFilename, $sFuncName, 1)
			RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveDirectory", "REG_SZ", $sAutosaveDirectory)
			RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveFilename", "REG_SZ", $sAutosaveFilename)
			RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveStartStandardProgram", "REG_SZ", $g_sProgresInjecteursAutoOpenOutPutFile)
			; On bascule sur PDFCreator
			$i = 0
			_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
			While _YDTool_GetDefaultPrinter(@ComputerName) <> $g_sPdfCreatorPrinter
				$i += 1
				Sleep(100)
				_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
				If $i > 20 Then
					_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule impossible vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
					Return False
				EndIf
			WEnd
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
			; On reactive PROGRES
			_YDTool_SuspendProcessSwitch($g_sProgresExeFileName, False)
			; L'impression se lance ...
			Sleep(1000)
			; On verifie si l'impression est terminee
			If _IsPrintStopFromInjLogFile() Then
				_YDLogger_Log("Impression injecteur terminée !", $sFuncName)
				; On retourne sur l'imprimante par defaut
				_YDTool_SetDefaultPrinter($g_sDefaultPrinter)
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Retour sur imprimante : " & $g_sDefaultPrinterName, 5000)
			EndIf
		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement des impressions d'ouvertures de journée PROGRES
; Syntax ........: _ModuleOuvertures()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 09/12/2019
; Notes .........:
;================================================================================================================================
Func _ModuleOuvertures()
	Local $sFuncName = "_ModuleOuvertures"
	; On verifie le contexte
	_CheckContext()
	; On ne travaille que si PROGRES est lance et si la fenetre d'ouverture de la journée PROGRES est détectée
	If ProcessExists($g_sProgresExeFileName) And WinActive($g_sProgresOuvertureWindowTitle) Then
		_YDLogger_Log("Fenetre ouverture PROGRES : ouverte", $sFuncName)
		; On recupere l'UGE
		Local $sTechUge = _GetUGEFromTechLogFile()
		; On modifie le registre pour modifier le Path et le nom du fichier
		$g_sSiteNetworkPath = ($g_sSite = "ARRAS") ? $g_sProgresOuvertureArrasPath : $g_sProgresOuvertureLensPath
		_YDLogger_Var("$g_sSiteNetworkPath", $g_sSiteNetworkPath)
		Local $sAutosaveDirectory = $g_sSiteNetworkPath & "\" & @YEAR & @MON & @MDAY & "\"
		Local $sAutosaveFilename = "<Datetime>_" & $sTechUge
		_YDLogger_Log("Modification du registre", $sFuncName, 1)
		_YDLogger_Var("$sAutosaveFilename", $sAutosaveFilename, $sFuncName, 1)
		_YDLogger_Var("$sAutosaveDirectory", $sAutosaveDirectory, $sFuncName, 1)
		RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveDirectory", "REG_SZ", $sAutosaveDirectory)
		RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveFilename", "REG_SZ", $sAutosaveFilename)
		RegWrite("HKEY_CURRENT_USER\Software\PDFCreator\Profiles\" & $g_sPdfCreatorPrinter & "\Program", "AutosaveStartStandardProgram", "REG_SZ", "0")
		; On bascule sur PDFCreator
		Local $i = 0
		_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
		While _YDTool_GetDefaultPrinter(@ComputerName) <> $g_sPdfCreatorPrinter
			$i += 1
			Sleep(100)
			_YDTool_SetDefaultPrinter($g_sPdfCreatorPrinter)
			If $i > 20 Then
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule impossible vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
				Return False
			EndIf
		WEnd
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Bascule vers imprimante : " & $g_sPdfCreatorPrinter, 5000)
		; On attend que la fenetre d'ouverture progres soit fermee
		WinWaitClose($g_sProgresOuvertureWindowTitle)
		_YDLogger_Log("Fenetre ouverture PROGRES : fermee", $sFuncName)
		; Tant que PROGRES existe, on verifie que l'UGE du NTIC_xxxx.LOG soit egale à la nouvelle UGE
		While ProcessExists($g_sProgresExeFileName)
			Local $sNticUge = _GetUGEFromNticLogFile()
			If $sNticUge = $sTechUge Then
				_YDLogger_Log("Bascule sur nouvelle UGE OK ! ", $sFuncName)
				ExitLoop
			EndIf
			Sleep(100)
		WEnd
		; On retourne sur l'imprimante par defaut
		_YDTool_SetDefaultPrinter($g_sDefaultPrinter)
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Retour sur imprimante : " & $g_sDefaultPrinterName, 5000)
		; On propose d'ouvrir le dossier
		If $g_sProgresOuvertureAutoOpenOutPutDir = 1 Then
			ShellExecute($sAutosaveDirectory)
		EndIf
;~ 		If ProcessExists($g_sProgresExeFileName) Then
;~ 			Local $iInputOpenFolder = MsgBox(4, _YDGVars_Get("sAppTitle"), "Souhaitez-vous ouvrir le dossier des liasses d'Ouverture de Journée PROGRES ?")
;~ 			If ($iInputOpenFolder = 6) Then
;~ 				ShellExecute($g_sSiteNetworkPath)
;~ 			Endif
;~ 		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si le debut de l'impression de liasse a été demandé via le DAT du NSREPORT
; Syntax ........: _IsPrintStartFromMcoDatFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStartFromMcoDatFile()
	Local $sFuncName = "_IsPrintStartFromMcoDatFile"
	Local $sFileDateTime = _ArrayToString(FileGetTime($g_sProgresLiassesMcoDatFilePath))
	$g_bProgresLiassesMcoDatFileChanged = False
	_YDLogger_Var("$sFileDateTime", $sFileDateTime, $sFuncName, 2)
	If $sFileDateTime <> $g_sProgresLiassesMcoDatFileDateTime Then
		$g_sProgresLiassesMcoDatFileDateTime = $sFileDateTime
		$g_bProgresLiassesMcoDatFileChanged = True
	Endif
	_YDLogger_Var("$g_sProgresLiassesMcoDatFileDateTime", $g_sProgresLiassesMcoDatFileDateTime, $sFuncName, 2)
	_YDLogger_Var("$g_bProgresLiassesMcoDatFileChanged", $g_bProgresLiassesMcoDatFileChanged, $sFuncName, 2)
	Return $g_bProgresLiassesMcoDatFileChanged
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si le debut de l'impression d'un état RELANCE a été demandé via le DAT du NSREPORT
; Syntax ........: _IsPrintStartFromEtatRelanceDatFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStartFromEtatRelanceDatFile()
	Local $sFuncName = "_IsPrintStartFromEtatRelanceDatFile"
	If FileExists($g_sProgresInjecteursEtatRelanceDatFilePath) Then
		Local $sFileDateTime = _ArrayToString(FileGetTime($g_sProgresInjecteursEtatRelanceDatFilePath))
		$g_bProgresInjecteursEtatRelanceDatFileChanged = False
		_YDLogger_Var("$sFileDateTime", $sFileDateTime, $sFuncName, 2)
		If $sFileDateTime <> $g_sProgresInjecteursEtatRelanceDatFileDateTime And FileGetSize($g_sProgresInjecteursEtatRelanceDatFilePath) > 0 Then
			$g_sProgresInjecteursEtatRelanceDatFileDateTime = $sFileDateTime
			$g_bProgresInjecteursEtatRelanceDatFileChanged = True
		Endif
	Else
		$g_bProgresInjecteursEtatRelanceDatFileChanged = False
	EndIf
	_YDLogger_Var("$g_sProgresInjecteursEtatRelanceDatFileDateTime", $g_sProgresInjecteursEtatRelanceDatFileDateTime, $sFuncName, 2)
	_YDLogger_Var("$g_bProgresInjecteursEtatRelanceDatFileChanged", $g_bProgresInjecteursEtatRelanceDatFileChanged, $sFuncName, 2)
	Return $g_bProgresInjecteursEtatRelanceDatFileChanged
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si le debut de l'impression d'un état REJET a été demandé via le DAT du NSREPORT
; Syntax ........: _IsPrintStartFromEtatRejetDatFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStartFromEtatRejetDatFile()
	Local $sFuncName = "_IsPrintStartFromEtatRejetDatFile"
	If FileExists($g_sProgresInjecteursEtatRejetDatFilePath) Then
		Local $sFileDateTime = _ArrayToString(FileGetTime($g_sProgresInjecteursEtatRejetDatFilePath))
		$g_bProgresInjecteursEtatRejetDatFileChanged = False
		_YDLogger_Var("$sFileDateTime", $sFileDateTime, $sFuncName, 2)
		If $sFileDateTime <> $g_sProgresInjecteursEtatRejetDatFileDateTime And FileGetSize($g_sProgresInjecteursEtatRejetDatFilePath) > 0 Then
			$g_sProgresInjecteursEtatRejetDatFileDateTime = $sFileDateTime
			$g_bProgresInjecteursEtatRejetDatFileChanged = True
		Endif
	Else
		$g_bProgresInjecteursEtatRejetDatFileChanged = False
	EndIf
	_YDLogger_Var("$g_sProgresInjecteursEtatRejetDatFileDateTime", $g_sProgresInjecteursEtatRejetDatFileDateTime, $sFuncName, 2)
	_YDLogger_Var("$g_bProgresInjecteursEtatRejetDatFileChanged", $g_bProgresInjecteursEtatRejetDatFileChanged, $sFuncName, 2)
	Return $g_bProgresInjecteursEtatRejetDatFileChanged
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si le debut de l'impression d'un état OK a été demandé via le DAT du NSREPORT
; Syntax ........: _IsPrintStartFromEtatOkDatFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStartFromEtatOkDatFile()
	Local $sFuncName = "_IsPrintStartFromEtatOkDatFile"
	If FileExists($g_sProgresInjecteursEtatOkDatFilePath) Then
		Local $sFileDateTime = _ArrayToString(FileGetTime($g_sProgresInjecteursEtatOkDatFilePath))
		$g_bProgresInjecteursEtatOkDatFileChanged = False
		_YDLogger_Var("$sFileDateTime", $sFileDateTime, $sFuncName, 2)
		If $sFileDateTime <> $g_sProgresInjecteursEtatOkDatFileDateTime And FileGetSize($g_sProgresInjecteursEtatOkDatFilePath) > 0 Then
			$g_sProgresInjecteursEtatOkDatFileDateTime = $sFileDateTime
			$g_bProgresInjecteursEtatOkDatFileChanged = True
		Endif
	Else
		$g_bProgresInjecteursEtatOkDatFileChanged = False
	EndIf
	_YDLogger_Var("$g_sProgresInjecteursEtatOkDatFileDateTime", $g_sProgresInjecteursEtatOkDatFileDateTime, $sFuncName, 2)
	_YDLogger_Var("$g_bProgresInjecteursEtatOkDatFileChanged", $g_bProgresInjecteursEtatOkDatFileChanged, $sFuncName, 2)
	Return $g_bProgresInjecteursEtatOkDatFileChanged
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si le debut de l'impression d'un état AV a été demandé via le DAT du NSREPORT
; Syntax ........: _IsPrintStartFromEtatAvDatFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStartFromEtatAvDatFile()
	Local $sFuncName = "_IsPrintStartFromEtatAvDatFile"
	If FileExists($g_sProgresInjecteursEtatAvDatFilePath) Then
		Local $sFileDateTime = _ArrayToString(FileGetTime($g_sProgresInjecteursEtatAvDatFilePath))
		$g_bProgresInjecteursEtatAvDatFileChanged = False
		_YDLogger_Var("$sFileDateTime", $sFileDateTime, $sFuncName, 2)
		If $sFileDateTime <> $g_sProgresInjecteursEtatAvDatFileDateTime And FileGetSize($g_sProgresInjecteursEtatAvDatFilePath) > 0 Then
			$g_sProgresInjecteursEtatAvDatFileDateTime = $sFileDateTime
			$g_bProgresInjecteursEtatAvDatFileChanged = True
		Endif
	Else
		$g_bProgresInjecteursEtatAvDatFileChanged = False
	EndIf
	_YDLogger_Var("$g_sProgresInjecteursEtatAvDatFileDateTime", $g_sProgresInjecteursEtatAvDatFileDateTime, $sFuncName, 2)
	_YDLogger_Var("$g_bProgresInjecteursEtatAvDatFileChanged", $g_bProgresInjecteursEtatAvDatFileChanged, $sFuncName, 2)
	Return $g_bProgresInjecteursEtatAvDatFileChanged
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si la fin de l'impression a été détectée dans le fichier TECH_xxxxxxx.LOG
; Syntax ........: _IsPrintStopFromTechLogFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 04/12/2019
; Notes .........:
;================================================================================================================================
Func _IsPrintStopFromTechLogFile()
	Local $sFuncName = "_IsPrintStopFromTechLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "F_APHI_IMPRIME_BORDEREAU F_AROB_COMPTER_OCCURRENCES"
	Local $iPrintStopDetected = 0
	Local $bPrintStop = False
	;------------------------------
	$hLogFile = FileOpen($g_sProgresTechLogFilePath, 0)
	If $hLogFile = -1 Then
		_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
		Return False
	Endif
	; On recupere le nombre de lignes du fichier de log
	Local $iFileCountLine = _FileCountLines($g_sProgresTechLogFilePath)
	_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
	; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
	If $iFileCountLine < $g_iLastLineTechLogFile Then $g_iLastLineTechLogFile = 1
	_YDLogger_Var("$g_iLastLineTechLogFile (avant)", $g_iLastLineTechLogFile, $sFuncName, 2)
	; On boucle sur le fichier log
	For $i = $iFileCountLine to $g_iLastLineTechLogFile Step -1
		If $iFileCountLine = $g_iLastLineTechLogFile Then ExitLoop
		$sFileLine = FileReadLine($hLogFile, $i)
		$iPrintStopDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
		; Si pattern trouve on sort de la boucle
		If $iPrintStopDetected > 0 Then
			;------------------------------
			_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
			$bPrintStop = True
			ExitLoop
		EndIf
	Next
	FileClose($hLogFile)
	;------------------------------
	; On log les infos utiles
	$g_iLastLineTechLogFile = $iFileCountLine
	_YDLogger_Var("$g_iLastLineTechLogFile (apres)", $g_iLastLineTechLogFile, $sFuncName, 2)
	Return $bPrintStop
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de verifier si la fin de l'impression a été détectée dans le fichier INJ_xxxxxxx.LOG
; Syntax ........: _IsPrintStopFromInjLogFile()
; Parameters ....:
; Return values .: True / False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 02/01/2020
; Notes .........:
;================================================================================================================================
Func _IsPrintStopFromInjLogFile()
	Local $sFuncName = "_IsPrintStopFromInjLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "Fin du traitement du deblocage detail"
	Local $iPrintStopDetected = 0
	Local $bPrintStop = False
	;------------------------------
	$hLogFile = FileOpen($g_sProgresInjLogFilePath, 0)
	If $hLogFile = -1 Then
		_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
		Return False
	Endif
	; On recupere le nombre de lignes du fichier de log
	Local $iFileCountLine = _FileCountLines($g_sProgresInjLogFilePath)
	_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
	; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
	If $iFileCountLine < $g_iLastLineInjLogFile Then $g_iLastLineInjLogFile = 1
	_YDLogger_Var("$g_iLastLineInjLogFile (avant)", $g_iLastLineInjLogFile, $sFuncName, 2)
	; On boucle sur le fichier log
	For $i = $iFileCountLine to $g_iLastLineInjLogFile Step -1
		If $iFileCountLine = $g_iLastLineInjLogFile Then ExitLoop
		$sFileLine = FileReadLine($hLogFile, $i)
		$iPrintStopDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
		; Si pattern trouve on sort de la boucle
		If $iPrintStopDetected > 0 Then
			;------------------------------
			_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
			$bPrintStop = True
			ExitLoop
		EndIf
	Next
	FileClose($hLogFile)
	;------------------------------
	; On log les infos utiles
	$g_iLastLineInjLogFile = $iFileCountLine
	_YDLogger_Var("$g_iLastLineInjLogFile (apres)", $g_iLastLineInjLogFile, $sFuncName, 2)
	Return $bPrintStop
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de récupérer l'UGE via le fichier TECH_xxxxxxx.LOG
; Syntax ........: _GetUGEFromTechLogFile()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 15/05/2019
; Notes .........:
;================================================================================================================================
Func _GetUGEFromTechLogFile()
	Local $sFuncName = "_GetUGEFromTechLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "F_LNPG_USER0 Appel du centre : "
	Local $sUGE = "0000"
	; On ne fait des recherches que si la fenetre d'ouverture de la journée PROGRES est détectée
	If WinExists($g_sProgresOuvertureWindowTitle) Then
		$hLogFile = FileOpen($g_sProgresTechLogFilePath, 0)
		If $hLogFile = -1 Then
			_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
			Return False
		Endif
		; On recupere le nombre de lignes du fichier de log
		Local $iFileCountLine = _FileCountLines($g_sProgresTechLogFilePath)
		_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
		; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
		If $iFileCountLine < $g_iLastLineTechLogFile Then $g_iLastLineTechLogFile = 1
		_YDLogger_Var("$g_iLastLineTechLogFile (avant)", $g_iLastLineTechLogFile, $sFuncName, 2)
		; On boucle sur le fichier log
		For $i = $iFileCountLine to $g_iLastLineTechLogFile Step -1
			$sFileLine = FileReadLine($hLogFile, $i)
			Local $iUGEDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
			; Si pattern trouve on sort de la boucle
			If $iUGEDetected > 0 Then
				;------------------------------
				_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
				$sUGE = StringRight($sFileLine, 4)
				$g_iLastLineTechLogFile = _FileCountLines($g_sProgresTechLogFilePath) + 1
				ExitLoop
			EndIf
		Next
		FileClose($hLogFile)
		;------------------------------
		; On log les infos utiles
		_YDLogger_Var("$sUGE", $sUGE, $sFuncName, 2)
		_YDLogger_Var("$g_iLastLineTechLogFile (apres)", $g_iLastLineTechLogFile, $sFuncName, 2)
		;------------------------------
		Return $sUGE
	Endif
	Return $sUGE
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de récupérer l'UGE via le fichier NTIC_xxxxxxx.LOG
; Syntax ........: _GetUGEFromNticLogFile()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 15/05/2019
; Notes .........:
;================================================================================================================================
Func _GetUGEFromNticLogFile()
	Local $sFuncName = "_GetUGEFromNticLogFile"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "Sz_UgeNum : "
	Local $sUGE = "0000"
	; On ne fait des recherches que si PROGRES est lance
	If ProcessExists($g_sProgresExeFileName) Then
		If FileExists($g_sProgresNticLogFilePath) = 0 Then
			_YDLogger_Log("Fichier non present : " & $g_sProgresNticLogFilePath, $sFuncName)
			Return False
		EndIf
		$hLogFile = FileOpen($g_sProgresNticLogFilePath, 0)
		If $hLogFile = -1 Then
			_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier de log inaccessible !", 0, $TIP_ICONASTERISK)
			Return False
		Endif
		; On recupere le nombre de lignes du fichier de log
		Local $iFileCountLine = _FileCountLines($g_sProgresNticLogFilePath)
		_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
		; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
		If $iFileCountLine < $g_iLastLineNticLogFile Then $g_iLastLineNticLogFile = 1
		_YDLogger_Var("$g_iLastLineNticLogFile (avant)", $g_iLastLineNticLogFile, $sFuncName, 2)
		; On boucle sur le fichier log
		For $i = $iFileCountLine to $g_iLastLineNticLogFile Step -1
			$sFileLine = FileReadLine($hLogFile, $i)
			Local $iUGEDetected = StringInStr($sFileLine, $sPattern, 0, 1, 1)
			; Si pattern trouve on sort de la boucle
			If $iUGEDetected > 0 Then
				;------------------------------
				_YDLogger_Log("Pattern trouve : " & $sPattern, $sFuncName, 2)
				$sUGE = StringRight($sFileLine, 4)
				$g_iLastLineNticLogFile = _FileCountLines($g_sProgresNticLogFilePath) + 1
				ExitLoop
			EndIf
		Next
		FileClose($hLogFile)
		;------------------------------
		; On log les infos utiles
		_YDLogger_Var("$sUGE", $sUGE, $sFuncName)
		_YDLogger_Var("$g_iLastLineTechLogFile (apres)", $g_iLastLineNticLogFile, $sFuncName, 2)
		;------------------------------
		Return $sUGE
	Endif
	Return $sUGE
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de vérifier si l'imprimante g_sPdfCreatorPrinter est bien installee
; Syntax.........: _IsPdfCreatorPrinterInstalled()
; Parameters ....:
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 24/05/2019
; Notes .........:
;================================================================================================================================
Func _IsPdfCreatorPrinterInstalled()
	Local $sFuncName = "_IsPdfCreatorPrinterInstalled"
	Local $sRegKey = "HKCU\Software\PDFCreator\Printers"
	Local $sRegVal = $g_sPdfCreatorPrinter
	Local $sRegKeyVal = $sRegKey & "\" & $sRegVal
	If RegRead($sRegKey, $sRegVal) <> "" Then
		;_YDLogger_Log("Cle [" & $sRegKeyVal & "] trouvee", $sFuncName, 2)
		Return True
	Else
		_YDLogger_Error("Cle [" & $sRegKeyVal & "] introuvable !", $sFuncName)
		Return False
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Restauration Imprimante + Reactivation forcee de Progres
; Syntax ........: _RestoreOnError()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 18/12/2019
; Notes .........:
;================================================================================================================================
Func _RestoreOnError()
	Local $sFuncName = "_RestoreOnError"
	If @error <> 0 Then
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Une anomalie a été détectée ! " & @CRLF & "Réactivation de PROGRES en cours ..." & @CRLF & "Retour sur imprimante : " & $g_sDefaultPrinterName, 5000, $TIP_ICONASTERISK)
		; On retourne sur l'imprimante par defaut
		_YDTool_SetDefaultPrinter($g_sDefaultPrinter)
		; On retourne sur l'imprimante par defaut
		If Not _YDTool_SuspendProcessSwitch($g_sProgresExeFileName, False) Then
			_YDLogger_Error("Erreur lors de la reactivation de Progres !", $sFuncName)
		EndIf
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fermeture de " & _YDGVars_Get("sAppTitle") & " suite à une anomalie !", 5000, $TIP_ICONASTERISK)
	Endif
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet d'installer l'imprimante g_sPdfCreatorPrinter
; Syntax.........: _InstallPdfCreatorPrinter()
; Parameters ....:
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 24/05/2019
; Notes .........:
;================================================================================================================================
Func _InstallPdfCreatorPrinter()
	Local $sFuncName = "_InstallPdfCreatorPrinter"
	Local $iRegError
	Local $sRegName
	;---------------------------------------
	$sRegName = 'HKCU_add_printer'
	$iRegError = 0
	If RegWrite('HKCU\Software\PDFCreator\Printers', $g_sPdfCreatorPrinter, 'REG_SZ', $g_sPdfCreatorPrinter) <> 1 Then $iRegError += 1
	If $iRegError = 0 Then
		_YDLogger_Log("Inscriptions registre " & $sRegName & " : OK", $sFuncName)
	Else
		_YDLogger_Error("Inscriptions registre " & $sRegName & " : NOK !", $sFuncName)
	EndIf
	;---------------------------------------
	$sRegName = 'HKCU_configure_landscape'
	$iRegError = 0
	Local $RegData = '4400490050005f00500044004600430072006500610074006f00720000000000000000000000000000000000000000000000'
	$RegData &= '000000000000000000000000000001040205dc00c40253ef8101020009009a0b3408640001000f0058020200010058020300'
	$RegData &= '0100410034000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000100000000000000010000000200000001000000'
	$RegData &= '000000000000000000000000000000000000000050524956e230000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '00000000000000001800000000001027102710270000102700000000000000000000c4020000000000000000000000000000'
	$RegData &= '0000000000000000000003000000000000000000100050340300288804000000000000000000000001000000000000000000'
	$RegData &= '0000000000000000e7b14b4c0300000005000a00ff0000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000'
	$RegData &= '00000000000000000000000000000000000000000000000000000000'
	If RegWrite('HKCU\Printers\DevModePerUser', $g_sPdfCreatorPrinter, 'REG_BINARY', Binary('0x' & $RegData)) <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Printers\DevModes2', $g_sPdfCreatorPrinter, 'REG_BINARY', Binary('0x' & $RegData)) <> 1 Then $iRegError += 1
	If $iRegError = 0 Then
		_YDLogger_Log("Inscriptions registre " & $sRegName & " : OK", $sFuncName)
	Else
		_YDLogger_Error("Inscriptions registre " & $sRegName & " : NOK !", $sFuncName)
	EndIf
	;---------------------------------------
	$sRegName = 'HKCU_add_profile'
	$iRegError = 0
	Local $sRegData = 'Microsoft Word - |\.docx|\.doc|\Microsoft Excel - |\.xlsx|\.xls|\Microsoft PowerPoint - |\.pptx|\.ppt|'
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Ghostscript', 'DirectoryGhostscriptBinaries','REG_SZ','C:\Program Files\PDFCreator\GS9.05\gs9.05\Bin\') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Ghostscript', 'DirectoryGhostscriptFonts','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Ghostscript', 'DirectoryGhostscriptLibraries','REG_SZ','C:\Program Files\PDFCreator\GS9.05\gs9.05\Lib\') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Ghostscript', 'DirectoryGhostscriptResource','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'Counter','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'DeviceHeightPoints','REG_SZ','157') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'DeviceWidthPoints','REG_SZ','222') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'OneFilePerPage','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'Papersize','REG_SZ','a4') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampFontColor','REG_SZ','#FF0000') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampFontname','REG_SZ','Arial') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampFontsize','REG_SZ','48') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampOutlineFontthickness','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampString','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StampUseOutlineFont','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardAuthor','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardCreationdate','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardDateformat','REG_SZ','YYYYMMDDHHNNSS') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardKeywords','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardMailDomain','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardModifydate','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardSaveformat','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardSubject','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'StandardTitle','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'UseCreationDateNow','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'UseCustomPaperSize','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'UseFixPapersize','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing', 'UseStandardAuthor','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'BMPColorscount','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'BMPResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'JPEGColorscount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'JPEGQuality','REG_SZ','75') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'JPEGResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PCLColorsCount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PCLResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PCXColorscount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PCXResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PNGColorscount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PNGResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PSDColorsCount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'PSDResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'RAWColorsCount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'RAWResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'SVGResolution','REG_SZ','72') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'TIFFColorscount','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\Bitmap\Colors', 'TIFFResolution','REG_SZ','150') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Colors', 'PDFColorsCMYKToRGB','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Colors', 'PDFColorsColorModel','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Colors', 'PDFColorsPreserveHalftone','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Colors', 'PDFColorsPreserveOverprint','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Colors', 'PDFColorsPreserveTransfer','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompression','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGHighFactor','REG_SZ','0.9') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGLowFactor','REG_SZ','0.25') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGManualFactor','REG_SZ','3') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGMaximumFactor','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGMediumFactor','REG_SZ','0.5') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorCompressionJPEGMinimumFactor','REG_SZ','0.1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorResample','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorResampleChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionColorResolution','REG_SZ','300') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompression','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGHighFactor','REG_SZ','0.9') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGLowFactor','REG_SZ','0.25') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGManualFactor','REG_SZ','3') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGMaximumFactor','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGMediumFactor','REG_SZ','0.5') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyCompressionJPEGMinimumFactor','REG_SZ','0.1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyResample','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyResampleChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionGreyResolution','REG_SZ','300') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionMonoCompression','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionMonoCompressionChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionMonoResample','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionMonoResampleChoice','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionMonoResolution','REG_SZ','1200') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Compression', 'PDFCompressionTextCompression','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Fonts', 'PDFFontsEmbedAll','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Fonts', 'PDFFontsSubSetFonts','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Fonts', 'PDFFontsSubSetFontsPercent','REG_SZ','100') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFGeneralASCII85','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFGeneralAutorotate','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFGeneralCompatibility','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFGeneralDefault','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFGeneralOverprint','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFOptimize','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFPageLayout','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFPageMode','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFStartPage','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\General', 'PDFUpdateMetadata','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFAes128Encryption','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFAllowAssembly','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFAllowDegradedPrinting','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFAllowFillIn','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFAllowScreenReaders','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFDisallowCopy','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFDisallowModifyAnnotations','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFDisallowModifyContents','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFDisallowPrinting','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFEncryptor','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFHighEncryption','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFLowEncryption','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFOwnerPass','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFOwnerPasswordString','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFUserPass','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFUserPasswordString','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Security', 'PDFUseSecurity','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningMultiSignature','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningPFXFile','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningPFXFilePassword','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureContact','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureLeftX','REG_SZ','100') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureLeftY','REG_SZ','100') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureLocation','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureOnPage','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureReason','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureRightX','REG_SZ','200') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureRightY','REG_SZ','200') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignatureVisible','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningSignPDF','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PDF\Signing', 'PDFSigningTimeServerUrl','REG_SZ','http://timestamp.globalsign.com/scripts/timstamp.dll') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PS\LanguageLevel', 'EPSLanguageLevel','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Printing\Formats\PS\LanguageLevel', 'PSLanguageLevel','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AdditionalGhostscriptParameters','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AdditionalGhostscriptSearchpath','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AddWindowsFontpath','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AllowSpecialGSCharsInFilenames','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AutosaveDirectory','REG_SZ', '') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AutosaveFilename','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AutosaveFormat','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'AutosaveStartStandardProgram','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ClientComputerResolveIPAddress','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'DisableEmail','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'DisableUpdateCheck','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'DontUseDocumentSettings','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'EditWithPDFArchitect','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'FilenameSubstitutions','REG_SZ', $sRegData) <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'FilenameSubstitutionsOnlyInTitle','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'Language','REG_SZ','french') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'LastSaveDirectory','REG_SZ','<MyFiles>\') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'LastUpdateCheck','REG_SZ','20190509') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'Logging','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'LogLines','REG_SZ','100') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'MaximumCountOfPDFArchitectToolTip','REG_SZ','5') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'NoConfirmMessageSwitchingDefaultprinter','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'NoProcessingAtStartup','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'NoPSCheck','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'OpenOutputFile','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'OptionsDesign','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'OptionsEnabled','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'OptionsVisible','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSaving','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingBitsPerPixel','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingDuplex','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingMaxResolution','REG_SZ','600') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingMaxResolutionEnabled','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingNoCancel','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingPrinter','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingQueryUser','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrintAfterSavingTumble','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'PrinterStop','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ProcessPriority','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ProgramFont','REG_SZ','MS Sans Serif') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ProgramFontCharset','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ProgramFontSize','REG_SZ','8') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RemoveAllKnownFileExtensions','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RemoveSpaces','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramAfterSaving','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramAfterSavingProgramname','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramAfterSavingProgramParameters','REG_SZ','"<OutputFilename>"') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramAfterSavingWaitUntilReady','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramAfterSavingWindowstyle','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramBeforeSaving','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramBeforeSavingProgramname','REG_SZ','') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramBeforeSavingProgramParameters','REG_SZ','"<TempFilename>"') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'RunProgramBeforeSavingWindowstyle','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'SaveFilename','REG_SZ','<Title>') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'SendEmailAfterAutoSaving','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'SendMailMethod','REG_SZ','0') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'ShowAnimation','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'Toolbars','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'UpdateInterval','REG_SZ','2') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'UseAutosave','REG_SZ','1') <> 1 Then $iRegError += 1
	If RegWrite('HKCU\Software\PDFCreator\Profiles\' & $g_sPdfCreatorPrinter & '\Program', 'UseAutosaveDirectory','REG_SZ','1') <> 1 Then $iRegError += 1
	If $iRegError = 0 Then
		_YDLogger_Log("Inscriptions registre " & $sRegName & " : OK", $sFuncName)
	Else
		_YDLogger_Error("Inscriptions registre " & $sRegName & " : NOK !", $sFuncName)
		Return False
	EndIf
EndFunc
