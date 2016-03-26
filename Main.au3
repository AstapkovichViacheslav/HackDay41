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

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <install.au3>

#include ".\Models\_sourceModel_XML.au3"	;Реализация для XML
_DebugSetup("DEBUG",true,2)					;Указываем режим дебага (2 = консоль)
Opt("MustDeclareVars", 1)					;Все переменные должны быть объявлены

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

;Запрос создание по файлу шаблона
Func MakeTemplate($Fpath)
	;Запускаем на анализ полученный путь
	Template_analyze($FPath)
EndFunc


Func _ErrFunc()
	;ConsoleWrite("!> ERR = "& $oError.description & @CRLF)
EndFunc


