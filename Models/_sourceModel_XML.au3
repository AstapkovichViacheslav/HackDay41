#include-once

Func _Source_XML()
    Local $oClass = _AutoItObject_Class()
    With $oClass
		.AddMethod("getType", 	"_SourceXML_getType")
		.AddMethod("setSrc", 	"_SourceXML_setSrc")
		.AddMethod("getSrc", 	"_SourceXML_getSrc")

        .AddMethod("create", 	"_SourceXML_create")
		.AddMethod("open", 		"_SourceXML_open")
		.AddMethod("isReady",	"_SourceXML_isReady")
		;.AddMethod("createNode","_Source_createNode")
		;.AddMethod("selectNode","_Source_selectNode")
		;.AddMethod("getData",	"_Source_getData");

        .AddProperty("obj", $ELSCOPE_PRIVATE, ObjCreate("Microsoft.XMLDOM") )
		.AddProperty("src", $ELSCOPE_PRIVATE, "")
    EndWith
	return $oClass.Object
EndFunc

;Проверка, что это модуль XML
Func _SourceXML_getType($oSelf)
	#forceref $oSelf
	return "_sourceModel"
EndFunc

;Указать путь к источнику
Func _SourceXML_setSrc($oSelf, $Path)
	#forceref $oSelf
	$oSelf.src = $Path
EndFunc

;Получить путь к источнику
Func _SourceXML_getSrc($oSelf, $Path)
	#forceref $oSelf
	return $oSelf.src
EndFunc

;Указать модель работы с данными
Func _SourceXML_create($oSelf)
	#forceref $oSelf
	Local $code = $oSelf.obj.save($oSelf.src)
	return FileExists($oSelf.src)
EndFunc

;Готов ли источник данных
Func _SourceXML_isReady($oSelf)
	return FileExists($oSelf.src)
EndFunc

;Открытие источника
Func _SourceXML_open($oSelf)
	if $oSelf.src = "" then
		ConsoleWrite("!> XML: Не указан путь" & @CRLF)
		return 0
	EndIf
	Local $code = $oSelf.obj.Load($oSelf.src)
	msgbox(0,"test",$oSelf.obj.readyState)
	;https://msdn.microsoft.com/en-us/library/ms753702(v=vs.85).aspx
	return $oSelf.obj.readyState = 4
EndFunc