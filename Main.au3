#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include-once
#RequireAdmin
#include <Debug.au3>									;���������� ��� ������
Global const $pref = ObjCreate("Scripting.Dictionary")	;���������� ���������
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")	;����� ������
_DebugSetup("DEBUG",true,2)								;��������� �������� (2 - �������)
#include ".\Lib\AutoItObject.au3"						;��������� ������� � ��������
#include "sourceModel.au3"								;DAO ���������
#include ".\Models\_sourceModel_XML.au3"				;DAO ���������� ��� XML
;����������� ����������� ���������
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
;���������� UDF ��� ��������
#include "Template.au3"									;������ � ���������
Opt("MustDeclareVars", 1)								;��� ���������� ������ ���� ���������

#Region Initialization
;��������� AutoIT Object
_AutoItObject_StartUp()
;���� �� ������������� - ������ ���������
IF not @Compiled Then
	msgbox(64,"����������","������� ������ ���� ������������� ��� �������",3)
	;exit
EndIf
;���� ��� ������� ���������� - �������� �� ���������
If $CmdLine[0]=0 Then
	Install()
	;Exit
EndIf
;���� �� ������� �� ����� - ������� ��������� �� �������
Main_InitPref()
#EndRegion

;���������� ����, �� �������� �� �������
Local $File 	= Main_getFileName()
;���������� ������ - ��������
Local $Template = Template_TemplateEditor()



;Local $Test = FileOpenDialog("Choose file",@ScriptDir,"TEst (*.*)")
;MakeTemplate($Test)

;Запрос создание по файлу шаблона
Func MakeTemplate($Fpath)
	;TODO: Тут мы выбираем шаблон

	;Подготавливаем данные шаблона
	Local $template_data = Template_prepare($FPath)
EndFunc




;������� ��������� ��������
Func Main_InitPref()
	$Pref.add("TemplatePath", @ScriptDir&"\Templates")
EndFunc

Func _ErrFunc()
	;ConsoleWrite("!> ERR = "& $oError.description & @CRLF)
EndFunc

Func Main_getFileName()
	;���� ��� ��������
	return @ScriptDir&"\TestRailExport.xml"
EndFunc