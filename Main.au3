#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#RequireAdmin
#include-once
#include <AutoItObject.au3>					;Поддержка объектов
#include <Debug.au3>						;Отладка и запись
#include "Template.au3"						;Логика работы с шаблонами
#include "sourceModel.au3"					;DAO интерфейс
#include ".\Models\_sourceModel_XML.au3"	;Реализация для XML
_DebugSetup("DEBUG",true,2)					;Указываем режим дебага (2 = консоль)
Opt("MustDeclareVars", 1)					;Все переменные должны быть объявлены

#Region Initialization
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")
_AutoItObject_StartUp()
#cs
If @Compiled Then
	msgbox(64,"info","compiled")
	Exit
	CreateMenuItem("AutoPIM")
EndIf
#ce
#EndRegion

Local $Test = FileOpenDialog("Choose file",@ScriptDir,"TEst (*.*)")
MakeTemplate($Test)

;Запрос создание по файлу шаблона
Func MakeTemplate($Fpath)
	;Запускаем на анализ полученный путь
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


