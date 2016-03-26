#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#RequireAdmin
Global const $pref = ObjCreate("Scripting.Dictionary")	;Глобальные настройки
#include-once
#include <AutoItObject.au3>								;Поддержка объектов
#include <Debug.au3>									;Отладка и запись
#include "Template.au3"									;Логика работы с шаблонами
#include "sourceModel.au3"								;DAO интерфейс
#include <File.au3>										;Работа с файлами
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include ".\Models\_sourceModel_XML.au3"				;Реализация для XML

_DebugSetup("DEBUG",true,2)								;Указываем режим дебага (2 = консоль)
Opt("MustDeclareVars", 1)								;Все переменные должны быть объявлены

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
;Инициализация переменных
Main_InitPref()
#EndRegion

;Main_Template_Edit()
Main_TemplateEditor()

;Local $Test = FileOpenDialog("Choose file",@ScriptDir,"TEst (*.*)")
;MakeTemplate($Test)

;Запрос создание по файлу шаблона
Func MakeTemplate($Fpath)
	;TODO: Тут мы выбираем шаблон

	;Подготавливаем данные шаблона
	Local $template_data = Template_prepare($FPath)


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

;Инициализация настроек (чтение из реестра и т.п.)
Func Main_InitPref()
	$Pref.add("TemplatePath", @ScriptDir&"\Templates")
EndFunc



;Редактирование элемента шаблона
Func Main_Template_Edit()
	Local $Param, $Value
	Local $Form1 = GUICreate("Редактор", 264, 119, 192, 124)
	Local $Group2 = GUICtrlCreateGroup("Элемент шаблона", 8, 8, 249, 105)
	Local $Label1 = GUICtrlCreateLabel("Входной элемент", 16, 32, 92, 17)
	Local $Label2 = GUICtrlCreateLabel("Элемент документа", 16, 64, 105, 17)
	Local $Input = GUICtrlCreateInput("MileStone", 128, 24, 120, 21)
	Local $Output = GUICtrlCreateInput("Title", 128, 56, 121, 21)
	Local $Button1 = GUICtrlCreateButton("Создать", 16, 80, 75, 25)
	Local $Button2 = GUICtrlCreateButton("Отмена", 176, 80, 75, 25)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUISetState(@SW_SHOW)
	Local $nMsg
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case -3
				GuiDelete($Form1)
				ExitLoop
			Case $Button1
				$Param = GUICtrlRead($Input)
				$Value = GUICtrlRead($Output)
				ExitLoop
			Case $Button2
				GuiDelete($Form1)
				ExitLoop
		EndSwitch
	WEnd
	Local $Data = [$Param, $Value]
	return $Data
EndFunc


Func Main_TemplateEditor()
	Local $tmp = Template_getTemplateList()
	Local $Form1 = GUICreate("��������", 368, 248, 192, 124)
	Local $Group1 = GUICtrlCreateGroup("������ ��������", 8, 8, 353, 137)
	Local $List1 = GUICtrlCreateListView("������ ������ ��������", 16, 24, 249, 110)
	Local $Button1 = GUICtrlCreateButton("�������", 272, 24, 75, 25)
	Local $Button2 = GUICtrlCreateButton("�������", 272, 56, 75, 25)
	Local $Button3 = GUICtrlCreateButton("�������", 272, 88, 75, 25)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Local $Group2 = GUICtrlCreateGroup("������������ ��������� ��������", 8, 152, 353, 89)
	Local $Label1 = GUICtrlCreateLabel("������� �������", 16, 168, 91, 17)
	Local $Combo1 = GUICtrlCreateCombo("<>", 16, 184, 153, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	Local $Label2 = GUICtrlCreateLabel("�������� �������", 184, 168, 99, 17)
	Local $Input1 = GUICtrlCreateInput("<>", 184, 184, 161, 21)
	Local $Button4 = GUICtrlCreateButton("�����", 16, 208, 75, 25)
	Local $Button5 = GUICtrlCreateButton("�������", 96, 208, 75, 25)
	Local $Button6 = GUICtrlCreateButton("��������� ������", 184, 208, 163, 25)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	;������� ������ ��������� ��������
	If IsArray($tmp) Then
		For $i=1 to $tmp[0]
			ConsoleWrite($tmp[$i]& @CRLF)
			GUICtrlCreateListViewItem($tmp[$i],$List1)
		Next
	EndIf
	;
	Local $src = _Source()
	GUISetState(@SW_SHOW)
	LocaL $nMsg
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			;�������� ������ ��������
			Case $Button3
				Local $SelectedPath = _GUICtrlListView_GetSelectedIndices($List1)
				$SelectedPath 		= _GUICtrlListView_GetItemTextString($List1, Int($SelectedPath))
				FileDelete($SelectedPAth)
				_GUICtrlListView_DeleteAllItems($List1)
				$tmp = Template_getTemplateList()
				If IsArray($tmp) Then
					For $i=1 to $tmp[0]
						ConsoleWrite($tmp[$i]& @CRLF)
						GUICtrlCreateListViewItem($tmp[$i],$List1)
					Next
				EndIf
			;�������� ������ ��������
			Case $Button1
				Local $tmpName = InputBox("�������� �������","������� �������� �������","TemplateNAme")
				Template_createTemplateFile($tmpName)
				_GUICtrlListView_DeleteAllItems($List1)
				$tmp = Template_getTemplateList()
				If IsArray($tmp) Then
					For $i=1 to $tmp[0]
						ConsoleWrite($tmp[$i]& @CRLF)
						GUICtrlCreateListViewItem($tmp[$i],$List1)
					Next
				EndIf
			;�������� ����� ������� ���������
			Case $Button2
				GUICtrlSetData($Combo1,"")
				GUICtrlSetData($Input1,"")
				Local $src = _Source()
				Local $SelectedPath = _GUICtrlListView_GetSelectedIndices($List1)
				$SelectedPath 		= _GUICtrlListView_GetItemTextString($List1, Int($SelectedPath))
				$src.setSrc($SelectedPath)
				$src.setModel( _Source_XML() )	;�������� ��� ������ � ����������: XML
				$src.open()
				Local $items = $src.getData()	;�������� ������ �� ���������
				_DebugOut("+> ��������� ��������: "& UBound($items)-1)
				For $i=0 to UBound($items)-1
					Local $Item = $items[$i][1] ;���� ��������
					If not IsArray($Item) then ContinueLoop
					GuiCtrlSetData($Combo1,$Item[2][1])
					ConsoleWrite($Item[2][1] & @CRLF)
				Next
			;�������� ��������
			Case $Combo1
				Local $Node = $src.select("Templates/Property[@name='"& GUICtrlRead($Combo1) &"']")
				GUICtrlSetData($Input1,$Node[3][1])
			;������� ��������
			Case $Button5
				Local $Node = $src.selectNode("./Property[@name='"& GUICtrlRead($Combo1) &"']")
				$src.deleteNode()
				GUICtrlSetData($Combo1,"")
				GUICtrlSetData($Input1,"")
			;���������� ���������
			Case $Button4
				Local $Attr = GUICtrlRead($Combo1)
				Local $AttrVal = GUICtrlRead($Input1)

				Local $Node = $src.selectNode("./Property[@name='"& $Attr &"']")
				If StringLen($Node.tagName) <> 0 then
					$src.setParam("value",$AttrVal)
					ContinueLoop
				EndIf

				$src.selectNode("")
				Local $Node = $src.element("Property")
				$Node.add("name",$Attr)
				$Node.add("value",GUICtrlRead($Input1))

				$src.createNode($Node)
				$src.selectNode("..")
				$src.save()
				GUICtrlSetData($Combo1,"")
			Case $Button6
				ExitLoop
		EndSwitch
	WEnd
	;������� ������ ������
	Local $SelectedPath = _GUICtrlListView_GetSelectedIndices($List1)
	$SelectedPath 		= _GUICtrlListView_GetItemTextString($List1, Int($SelectedPath))
	$src.setSrc($SelectedPath)
	$src.setModel( _Source_XML() )	;�������� ��� ������ � ����������: XML
	$src.open()
	Local $items = $src.getData()	;�������� ������ �� ���������
	;��������� ����
	GUIDelete($Form1)
	return $items
EndFunc