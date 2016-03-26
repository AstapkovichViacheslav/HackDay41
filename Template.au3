#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <Debug.au3>
#include "OOoCOM_UDF_v08.au3"
_DebugSetup("test",false,2)
; Script Start - Add your code below here
;Template_analyze("C:\WORK\HackDay41\Templates\default_templ.doc")



;Проверка: Установлен ли ворд на машине
Func Template_isWordInstalled()
	Local $Test = ObjCreate("Word.Application")
	Local $check = IsObj($Test)
	$Test = ""
	return $Check
EndFunc

;Сохранение ворд документа как XML (только Microsoft Word)
Func Template_saveDocAsXMLwithWord($FilePath)
	Local $fName 	= StringMid($FilePath,1, StringLen($FilePath)-3)&"xml"
	Local $obj 		= ObjCreate("Word.Application")
	If IsObj($obj) Then
		$obj.Documents.Open($FilePath)
		$obj.ActiveDocument.SaveAs($fName,11)
		$obj.ActiveDocument.Close
		$obj.quit()
	EndIf
	return $fName
EndFunc

Func Template_getTemplateList()
	Local $PathStr = _FileListToArray( $Pref("TemplatePath"), "tmp_*.xml")
	If not IsArray($PathStr) then return ""
	For $i=1 to $PathStr[0]
		$PathStr[$i]=  $Pref("TemplatePath") &"\"& $PathStr[$i]
	Next
	return $PathStr
EndFunc

Func Template_createTemplateFile($FileName)
	If $FileName = "" then return ""
	Local $Path = $Pref("TemplatePath") &"\tmp_"& $FileName & ".xml"
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc($Path)				;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.create()
	$src.open()
	Local $Node = $src.element("Templates")
	$src.createNode($Node)
	$src.save()
	return $src
EndFunc

Func Template_prepare($FilePath)
	IF not FileExists($FilePath) Then
		_DebugOut("!> Для шаблона передан несуществующий файл")
		return 0
	EndIf
	Local $Path = $FilePath
	Local $type = StringMid($Path,StringLen($Path)-2,3)
	If $type = "doc" Then
		If Template_isWordInstalled() Then
			$Path = Template_saveDocAsXMLwithWord($Path)
			$type = StringMid($Path,StringLen($Path)-2,3)
			_DebugOut("+> Сохранили doc в xml: "&$Path)
		Else
			_DebugOut("!> Microsoft Word не установлен в системе. Экспорт из doc в xml недоступен!")
			return
		EndIf
	EndIf
	;Проверим, что на данном этапе у нас XML
	If $type <> "xml" Then
		_DebugOut("!> Для анализа передан файл в формате, отличном от xml!")
		return
	EndIf
	;Теперь, будем искать нужные нам узлы
	msgbox(0,"Run analyze","ok")
	$src = _Source()				;Создаю ресурс
	$src.setSrc($Path)				;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.open()						;Открываю источник
	Local $dataSet = $src.getData()	;Получаю данные
	return $dataSet
EndFunc