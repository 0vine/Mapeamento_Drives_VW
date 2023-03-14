#cs ----------------------------------------------------------------------------

 AutoIt Version: 
 Author:     Vinicius Moreira

 Script Function:
	Mapear Drives
	

#ce ----------------------------------------------------------------------------

;#RequireAdmin

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPI.au3>
#include <GuiStatusBar.au3>
#include <ProgressConstants.au3>
#include <Misc.au3>
#include <Inet.au3>
#include <GDIPlus.au3>
#include <MsgBoxConstants.au3>
#include <GifAnimation.au3>
#include <UDF\UDF_Embed.au3>
#include <UDF\UDF_AutoitObject.au3>

;Variaveis para sombra
Global $hWnd_ShadowApp, $hImage, $hImageShadow
;Mensagem do popup
Global $sText = ""

Global $sLabel_Status, $idBtn_Execute, $Progress1, $StatusBar1
Global $sName = 'Interface'
Global $sFolder = 'C:\ProgramData\' & $sName
Global $sFileConfig = $sFolder & '\config.ini'

;Evita abrir mais de uma instancia do programa
If _Singleton(@ScriptName, 1) = 0 Then Exit

$hWnd_Main = GUICreate("ManagedBy", 250, 270, -1, -1, $WS_POPUP)
GUISetFont(12, 400, 0, "Segoe UI")

;Coleta a posicao da janela
$aPos = WinGetPos($hWnd_Main)
$g_iWidth = $aPos[2]
$g_iHeight = $aPos[3]

$sLabel_Status = GUICtrlCreateLabel("", 50, 140, 300, 50)
$Progress1 = GUICtrlCreateProgress(30, 120, 200, 10)
DllCall("UxTheme.dll", "int", "SetWindowTheme", "hwnd", GUICtrlGetHandle($Progress1), "wstr", "", "wstr", "")
GUICtrlSetState(-1, $GUI_HIDE)



;Cria borda arredondada
$hRgn = _WinAPI_CreateRoundRectRgn(0, 0, $g_iWidth, $g_iHeight, 20, 20)
_WinAPI_SetWindowRgn($hWnd_Main, $hRgn)

$idTextMap = GUICtrlCreateLabel("Mapeamento Drives", 50, 60)

$idImage_Logo_Volkswagen = GUICtrlCreatePic("", 95, 16, 50, 50)
_GUICtrlSetGIF(-1, _Extract_Logo_Volkswagen(True))

$idBtn_Executar = GUICtrlCreateButton("Executar", 65, 165,120, 40)

$idBtn_Postpone = GUICtrlCreateButton("Sair", 80, 215, 90, 30)
;$idBtn_Send = GUICtrlCreateButton("Enviar", 400, 400, 115, 41)

;Efetua a exibicao da janela com efeito
_WinAPI_AnimateWindow($hWnd_Main, Struct_AnimateList().fadeIn, 300)
GUISetState(@SW_SHOW)

;Exibe a sombra da janela
Void_DisplayShadowGUI()


;Laço repetição para GPupdate

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $idBtn_Postpone 
			Exit
        
		Case $idBtn_Executar
			GUICtrlSetState($Progress1, $GUI_SHOW)
			GUICtrlSetState($idBtn_Execute, $GUI_DISABLE)
			GUICtrlSetData($sLabel_Status, 'Mapeando Drives ..')
			_SendMessage(GUICtrlGetHandle($Progress1), $PBM_SETMARQUEE, True, 32)
			Sleep(2000)

			GUICtrlSetStyle($Progress1, $PBM_SETMARQUEE)
			_SendMessage(GUICtrlGetHandle($Progress1), $PBM_SETMARQUEE, True, 32)
			Local $sGPO_Auto = IniRead($sFileConfig, 'GPO', 'AutoRefresh', 'True')
			If $sGPO_Auto = 'True' Then 
				GUICtrlSetData($sLabel_Status, 'Aguarde Mapeando ..')
				RunWait('gpupdate /force', '', @SW_HIDE)
				Sleep(2000)
			EndIf

			GUICtrlSetData($sLabel_Status, '')
			GUICtrlSetState($Progress1, $GUI_HIDE)
			GUICtrlSetState($idBtn_Execute, $GUI_ENABLE)

	EndSwitch
WEnd




;Efetua a saida do programa de forma adequada
Func Void_Exit($Param = 'Exit')
	Local $hWnd_Exit, $hRgn2

	GUIDelete($hWnd_ShadowApp)
	
	If $Param = 'Exit' Then
		$hWnd_Exit = GUICreate("", 930, 680, 3, 3, $WS_POPUP, $WS_EX_MDICHILD, $hWnd_Main)
		GUISetFont(15, 400, 0, 'Segoe UI')
		GUISetBkColor(Struct_Colors().blue.light)
		$hRgn2 = _WinAPI_CreateRoundRectRgn(0, 0, $g_iWidth, $g_iHeight, 20, 20)
		_WinAPI_SetWindowRgn($hWnd_Exit, $hRgn2)
		
		GUICtrlCreateLabel(StringFormat('Aguarde...%sEnviando formulário', @CRLF), 0, $aPos[3] / 2, $aPos[2] - 1, 90, $SS_CENTER)
		GUICtrlSetBkColor(-1, Struct_Colors().alpha)
		GUICtrlSetColor(-1, Struct_Colors().white.normal)

		GUICtrlCreatePic('', 390, 610, 141, 30)
		_GUICtrlSetGIF(-1, _Extract_Logo_Stefanini(True))
		
		_WinAPI_AnimateWindow($hWnd_Exit, Struct_AnimateList().explode, 300)
		GUISetState()
		Sleep(3000)
	EndIf
	GUIDelete($hWnd_ShadowApp)
	_WinAPI_AnimateWindow($hWnd_Main, Struct_AnimateList().fadeOut, 300)
	_WinAPI_AnimateWindow($hWnd_Exit, Struct_AnimateList().fadeOut, 300)
	_WinAPI_DeleteObject($hRgn)
	_WinAPI_DeleteObject($hRgn2)
	_GDIPlus_ImageDispose($hImage)
	_GDIPlus_ImageDispose($hImageShadow)
	_GDIPlus_Shutdown()
	Exit
EndFunc

Func _Extract_FileConfig($bSaveBinary = False, $sSavePath = $sFolder)
	Local $Extract_FileConfig
	$Extract_FileConfig &= 'SbAAW0lQXQ0KUmUAbGVhc2U9RmGIbHNlAXBuZXcFYBBBdXRvBSxbR1AGTwC4AUhSZWZyZRxzaAZkASYANElQQ+BvbmZpZwUuAGAFFA=='
	$Extract_FileConfig = _WinAPI_Base64Decode($Extract_FileConfig)
	Local $tSource = DllStructCreate('byte[' & BinaryLen($Extract_FileConfig) & ']')
	DllStructSetData($tSource, 1, $Extract_FileConfig)
	Local $tDecompress
	_WinAPI_LZNTDecompress($tSource, $tDecompress, 107)
	$tSource = 0
	Local Const $bString = Binary(DllStructGetData($tDecompress, 1))
	If $bSaveBinary Then
		Local Const $hFile = FileOpen($sSavePath & "\config.ini", 18)
		If @error Then Return SetError(1, 0, 0)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return $bString
EndFunc   ;==>_Extract_FileConfig

;Lista de animacoes
Func Struct_AnimateList()
	Local Const $__t_STRUCT = 'struct;' & _
		'char fadeIn[10];' 				& _
		'char fadeOut[10];' 			& _
		'char explode[10];' 			& _
		'char implode[10];' 			& _
		'endstruct'
	
	Local $__AW = DllStructCreate($__t_STRUCT)

	DllStructSetData($__AW, 'fadeIn', 0x00080000)
	DllStructSetData($__AW, 'fadeOut', 0x00090000)
	DllStructSetData($__AW, 'explode', 0x00040010)
	DllStructSetData($__AW, 'implode', 0x00050010)

	Return $__AW
EndFunc

;Constroe uma sombra na janela principal
Func Void_DisplayShadowGUI()
	_GDIPlus_Startup()
	_Extract_Shadow(True)
	Dim $hWnd_ShadowApp = GUICreate("", 0, 0, -22, -17, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_TOOLWINDOW, $WS_EX_MDICHILD), $hWnd_Main)
	GUISetState(@SW_SHOW)
	Dim $hImage = _GDIPlus_ImageLoadFromFile(@TempDir & '\Shadow.png')
	Dim $hImageShadow = _GDIPlus_ImageResize($hImage, 305, 318)
	For $i = 0 To 255 Step 10
		Void_DrawPNG($i, $hWnd_ShadowApp, $hImageShadow)
	Next
	GUISwitch($hWnd_Main)
EndFunc   ;==>shadowGUI

;Desenha arquivos png
Func Void_DrawPNG($i, $sStrGui, $sStrSplashImage)
	Local $hScrDC, $hMemDC, $hBitmap, $hOld, $pSize, $tSize, $pSource, $tSource, $pBlend, $tBlend
	Local Const $AC_SRC_ALPHA = 1
	$hScrDC = _WinAPI_GetDC(0)
	$hMemDC = _WinAPI_CreateCompatibleDC($hScrDC)
	$hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($sStrSplashImage)
	$hOld = _WinAPI_SelectObject($hMemDC, $hBitmap)
	$tSize = DllStructCreate($tagSIZE)
	$pSize = DllStructGetPtr($tSize)
	DllStructSetData($tSize, "X", _GDIPlus_ImageGetWidth($sStrSplashImage))
	DllStructSetData($tSize, "Y", _GDIPlus_ImageGetHeight($sStrSplashImage))
	$tSource = DllStructCreate($tagPOINT)
	$pSource = DllStructGetPtr($tSource)
	$tBlend = DllStructCreate($tagBLENDFUNCTION)
	$pBlend = DllStructGetPtr($tBlend)
	DllStructSetData($tBlend, "Alpha", $i)
	DllStructSetData($tBlend, "Format", $AC_SRC_ALPHA)
	_WinAPI_UpdateLayeredWindow($sStrGui, $hScrDC, 0, $pSize, $hMemDC, $pSource, 0, $pBlend, $ULW_ALPHA)
	_WinAPI_ReleaseDC(0, $hScrDC)
	_WinAPI_SelectObject($hMemDC, $hOld)
	_WinAPI_DeleteObject($hBitmap)
	_WinAPI_DeleteDC($hMemDC)
	Sleep(5)
	GUISetState()
EndFunc   ;==>drawPNG



;Paleta de cores
Func Struct_Colors()
	Local $oColor = newObject()
	DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct')

	addObjectProperty($oColor, 'pink', 		DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'green', 	DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'yellow', 	DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'cyan', 		DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'white', 	DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'blue', 		DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'black', 	DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'red', 		DllStructCreate('struct;char normal[8];char light[8];char dark[8];endstruct'))
	addObjectProperty($oColor, 'alpha', 	-2)

	$oColor.pink.normal = 0xd900d9
	$oColor.pink.light = 0xff00ff
	$oColor.pink.dark = 0xd100d1

	$oColor.green.normal = 0x00d900
	$oColor.green.light = 0x00ff00
	$oColor.green.dark = 0x00bf00

	$oColor.yellow.normal = 0xe3e300
	$oColor.yellow.light = 0xffff00 
	$oColor.yellow.dark = 0xb5b500

	$oColor.cyan.normal = 0x00ffff
	$oColor.cyan.light = 0x00dbdb
	$oColor.cyan.dark = 0x00a8a8

	$oColor.white.normal = 0xeeeeee
	$oColor.white.light = 0xffffff
	$oColor.white.dark = 0xdddddd

	$oColor.blue.normal = 0x0000ff
	$oColor.blue.light = 0x125bad
	$oColor.blue.dark = 0x0011a8

	$oColor.black.normal = 0x000000
	$oColor.black.light = 0x353535
	$oColor.black.dark = 0x252525

	$oColor.red.normal = 0xd60000
	$oColor.red.light = 0xff0000
	$oColor.red.dark = 0xa80000

	Return $oColor
EndFunc

;Efeito ao passar o mouse ..

