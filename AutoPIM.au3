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
#include "Generator.au3"								;Работа с генератором документов
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

#Region Generation
;Определяем файл, по которому мы вызваны
Local $File 		= Main_getFileName()
;Определяем шаблон - развязку (получаем ресурс шаблона)
Local $TemplateSrc 	= Template_TemplateEditor()
;Получаю данные из файла источника данных(файл как ресурс)
Local $SourceFile 	= Main_getFileSource($File)
;Выбираем Файл оформления
Local $StyleFile 	= FileOpenDialog("Выбрать оформление", @ScriptDir,"All (*.xml)")
Local $StyleSrc		= Main_getFileSource($StyleFile)
;Генерируем на основе данных файл
;1 - Создание файла по шаблону
Generator_createDoc($StyleSrc, "C:\test-doc.xml")
#EndRegion
;Local $Gen = Main_Generate($ParsedData, $Template, $StyleFile)


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

;Полаем выбранный источник как ресурс
Func Main_getFileSource($FilePath)
	_DebugOut("+> getFilSource '"& $FilePath&"'"& @CRLF)
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc($FilePath)			;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.open()						;Открываю источник
	return $src
EndFunc

Func Main_getFileData($FilePath)
	_DebugOut("+> getFilSource '"& $FilePath&"'"& @CRLF)
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc($FilePath)			;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.open()						;Открываю источник
	return $src.getData()
EndFunc



Func Main_Generate($ParsedData, $Template, $StyleFile)
	Local $Data = $ParsedData.getData()
	Local $MileStoneName 	= $ParsedData.select("milestone/name",$Data)
	Local $TestPlanName		= $ParsedData.select("milestone/runs/plan/name",$Data)
	Local $TestPlanRuns 	= $ParsedData.select("milestone/runs/plan/runs",$Data)
	For $i=1 to UBound($TestPlanRuns)-1
		ConsoleWrite($TestPlanRuns[2][0] & @CRLF)
		Main_RunAnalyze($TestPlanRuns[2][1])
	Next
EndFunc

Func Main_RunAnalyze($TestRun)
	ConsoleWrite( UBound($TestRun)-1 & @CRLF)
	For $i=0 to UBound($TestRun)-1
		ConsoleWrite($TestRun[$i][0] & @CRLF)
	Next
EndFunc

