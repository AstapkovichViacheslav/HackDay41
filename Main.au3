#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include-once
#RequireAdmin
#include <Debug.au3>									;Библиотека для дебага
Global const $pref = ObjCreate("Scripting.Dictionary")	;Глобальные настройки
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")	;Ловим ошибки
_DebugSetup("DEBUG",true,2)								;Настройка дебагера (2 - консоль)
#include ".\Lib\AutoItObject.au3"						;Поддержка классов и объектов
#include "sourceModel.au3"								;DAO интерфейс
#include ".\Models\_sourceModel_XML.au3"				;DAO реализация для XML
;Подключение стандартных библиотек
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <install.au3>
;Подключаем UDF для шаблонов
#include "Template.au3"									;Работа с шаблонами
Opt("MustDeclareVars", 1)								;Все переменные должны быть объявлены

#Region Initialization
;Запускаем AutoIT Object
_AutoItObject_StartUp()
;Если не скомпилирован - нельзя выполнять
IF not @Compiled Then
	msgbox(64,"Информация","АвтоПИМ должен быть скомпилирован для запуска",3)
	;exit
EndIf
;Если нет входных параметров - запущены на установку
If $CmdLine[0]=0 Then
	Install()
	;Exit
EndIf
;Если мы запщены по файлу - считаем настройки из реестра
Main_InitPref()
#EndRegion

;Определяем файл, по которому мы вызваны
Local $File 	= Main_getFileName()
;Определяем шаблон - развязку
Local $Template = Template_TemplateEditor()



;Local $Test = FileOpenDialog("Choose file",@ScriptDir,"TEst (*.*)")
;MakeTemplate($Test)

;Р—Р°РїСЂРѕСЃ СЃРѕР·РґР°РЅРёРµ РїРѕ С„Р°Р№Р»Сѓ С€Р°Р±Р»РѕРЅР°
Func MakeTemplate($Fpath)
	;TODO: РўСѓС‚ РјС‹ РІС‹Р±РёСЂР°РµРј С€Р°Р±Р»РѕРЅ

	;РџРѕРґРіРѕС‚Р°РІР»РёРІР°РµРј РґР°РЅРЅС‹Рµ С€Р°Р±Р»РѕРЅР°
	Local $template_data = Template_prepare($FPath)
EndFunc




;Считать настройки АвтоПИМа
Func Main_InitPref()
	$Pref.add("TemplatePath", @ScriptDir&"\Templates")
EndFunc

Func _ErrFunc()
	;ConsoleWrite("!> ERR = "& $oError.description & @CRLF)
EndFunc

Func Main_getFileName()
	;Пока что заглушка
	return @ScriptDir&"\TestRailExport.xml"
EndFunc