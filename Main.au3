#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#RequireAdmin
#include-once
#include <AutoItObject.au3>					;РџРѕРґРґРµСЂР¶РєР° РѕР±СЉРµРєС‚РѕРІ
#include <Debug.au3>						;РћС‚Р»Р°РґРєР° Рё Р·Р°РїРёСЃСЊ
#include "Template.au3"						;Р›РѕРіРёРєР° СЂР°Р±РѕС‚С‹ СЃ С€Р°Р±Р»РѕРЅР°РјРё
#include "sourceModel.au3"					;DAO РёРЅС‚РµСЂС„РµР№СЃ

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

#include ".\Models\_sourceModel_XML.au3"	;Р РµР°Р»РёР·Р°С†РёСЏ РґР»СЏ XML
_DebugSetup("DEBUG",true,2)					;РЈРєР°Р·С‹РІР°РµРј СЂРµР¶РёРј РґРµР±Р°РіР° (2 = РєРѕРЅСЃРѕР»СЊ)
Opt("MustDeclareVars", 1)					;Р’СЃРµ РїРµСЂРµРјРµРЅРЅС‹Рµ РґРѕР»Р¶РЅС‹ Р±С‹С‚СЊ РѕР±СЉСЏРІР»РµРЅС‹

#Region Initialization
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")
_AutoItObject_StartUp()

If @Compiled Then
	Exit
	;CreateMenuItem("AutoPIM")
EndIf
#EndRegion

Install();

Local $Test = FileOpenDialog("Choose file",@ScriptDir,"TEst (*.*)")
MakeTemplate($Test)

;Р—Р°РїСЂРѕСЃ СЃРѕР·РґР°РЅРёРµ РїРѕ С„Р°Р№Р»Сѓ С€Р°Р±Р»РѕРЅР°
Func MakeTemplate($Fpath)
	;Р—Р°РїСѓСЃРєР°РµРј РЅР° Р°РЅР°Р»РёР· РїРѕР»СѓС‡РµРЅРЅС‹Р№ РїСѓС‚СЊ
	Template_analyze($FPath)


EndFunc

Func CreateMenuItem($GroupName)
	Local $Type = "xml"
	Local $Key = "HKEY_CLASSES_ROOT\."&$Type&"\Shell\"&@ScriptName
	RegWrite($key)
	RegWrite($key,"MUIVerb","REG_SZ",$GroupName)
	RegWrite($key,"SubCommands","REG_SZ",$GroupName&"1" &";"& $Groupname&"2")
	;http://rapidsoft.org/articles/wintuning/item/101-context_menu_section
EndFunc

Func _ErrFunc()
	;ConsoleWrite("!> ERR = "& $oError.description & @CRLF)
EndFunc

Func Install()
	Local $installStep = 1;

	local $InstallForm = GUICreate("Установка  АвтоПИМ", 501, 279, 413, 190)
	local $Logo = GUICtrlCreatePic(".\Images\АвтоПИМ.jpg", 0, 0, 500, 85)
	local $NextButton = GUICtrlCreateButton("Далее", 368, 248, 123, 25)
	local $Org_Label = GUICtrlCreateLabel("Организация", 43, 122, 71, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $Steps_Label = GUICtrlCreateLabel("", 24, 96, 449, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $F_Label = GUICtrlCreateLabel("Фамилия", 60, 153, 53, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $I_Label = GUICtrlCreateLabel("Имя", 85, 183, 26, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $O_Label = GUICtrlCreateLabel("Отчество", 62, 213, 51, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $Org_Input = GUICtrlCreateInput("", 120, 120, 250, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $F_Input = GUICtrlCreateInput("", 120, 150, 121, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $I_Input = GUICtrlCreateInput("", 120, 180, 121, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $O_Input = GUICtrlCreateInput("", 120, 210, 121, 21)
	GUICtrlSetState(-1, $GUI_HIDE)

	local $Text = GUICtrlCreateEdit("", 24, 96, 449, 129, BitOR($ES_READONLY,$ES_WANTRETURN), 0)
	GUICtrlSetData(-1, StringFormat("Вас приветствует инсталлятор программы АвтоПИМ\r\n\r\nАвтоПИМ упрощает процесс сборки документации из сторонних систем хранения.\r\nПозволяет гибко настраивать шаблоны и фильтры для создания документации.\r\n\r\nВ процессе установки программа АвтоПИМ проверит ранее установленные версии. \r\nПри наличии раннее установленных версий выполнит обновление. \r\n\r\nДля продолжения  установки нажмите кнопку "&Chr(34)&"Далее"&Chr(34)&""))

	GUISetState(@SW_SHOW)

	While 1
		local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $NextButton
				Switch $installStep
					Case 1
						$installStep = $installStep + 1
						GUICtrlSetData($Text, StringFormat("Шаг 1: Поиск ранее установленной версии АвтоПИМ"))
					Case 2
						$installStep = $installStep + 1
						GUICtrlSetState($Text, $GUI_HIDE)
						GUICtrlSetData($Steps_Label, StringFormat("Шаг 2: Регистрация данных о пользователе в системе"))
						GUICtrlSetState($Steps_Label, $GUI_SHOW)
						GUICtrlSetState($Org_Label, $GUI_SHOW)
						GUICtrlSetState($F_Label, $GUI_SHOW)
						GUICtrlSetState($I_Label, $GUI_SHOW)
						GUICtrlSetState($O_Label, $GUI_SHOW)
						GUICtrlSetState($Org_Input, $GUI_SHOW)
						GUICtrlSetState($F_Input, $GUI_SHOW)
						GUICtrlSetState($I_Input, $GUI_SHOW)
						GUICtrlSetState($O_Input, $GUI_SHOW)
				EndSwitch
		EndSwitch
	WEnd
EndFunc
