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

	local $InstallForm = GUICreate("���������  �������", 501, 279, 413, 190)
	local $Logo = GUICtrlCreatePic(".\Images\�������.jpg", 0, 0, 500, 85)
	local $NextButton = GUICtrlCreateButton("�����", 368, 248, 123, 25)
	local $Org_Label = GUICtrlCreateLabel("�����������", 43, 122, 71, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $Steps_Label = GUICtrlCreateLabel("", 24, 96, 449, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $F_Label = GUICtrlCreateLabel("�������", 60, 153, 53, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $I_Label = GUICtrlCreateLabel("���", 85, 183, 26, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	local $O_Label = GUICtrlCreateLabel("��������", 62, 213, 51, 17)
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
	GUICtrlSetData(-1, StringFormat("��� ������������ ����������� ��������� �������\r\n\r\n������� �������� ������� ������ ������������ �� ��������� ������ ��������.\r\n��������� ����� ����������� ������� � ������� ��� �������� ������������.\r\n\r\n� �������� ��������� ��������� ������� �������� ����� ������������� ������. \r\n��� ������� ������ ������������� ������ �������� ����������. \r\n\r\n��� �����������  ��������� ������� ������ "&Chr(34)&"�����"&Chr(34)&""))

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
						GUICtrlSetData($Text, StringFormat("��� 1: ����� ����� ������������� ������ �������"))
					Case 2
						$installStep = $installStep + 1
						GUICtrlSetState($Text, $GUI_HIDE)
						GUICtrlSetData($Steps_Label, StringFormat("��� 2: ����������� ������ � ������������ � �������"))
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
