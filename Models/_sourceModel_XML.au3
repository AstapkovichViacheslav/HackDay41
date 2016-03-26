#include-once
;https://msdn.microsoft.com/en-us/library/ms760218(v=vs.85).aspx

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
		.AddMethod("getData",	"_SourceXML_getData");

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
	#forceref $oSelf
	if $oSelf.src = "" then
		ConsoleWrite("!> XML: Не указан путь" & @CRLF)
		return 0
	EndIf
	Local $code = $oSelf.obj.Load($oSelf.src)
	;https://msdn.microsoft.com/en-us/library/ms753702(v=vs.85).aspx
	return $oSelf.obj.readyState = 4
EndFunc

Func _SourceXML_getData($oSelf, $oNode="")
	#forceref $oSelf
	Local $Node
	IF $oNode = "" then
		$Node = $oSelf.obj.documentElement
	Else
		$Node = $oNode
	EndIf
	IF StringLen($Node.tagName) = 0 then return ""
	Local $Attribs[2][2]
	;Стандартные атрибуты узла
	$Attribs[0][0] = "tagName"
	$Attribs[0][1] = $Node.tagName
	$Attribs[1][0] = "text"
	Local $RegExp = StringRegExp($Node.xml,"<\w",3)
	if UBound($RegExp)-1 = 0 then $Attribs[1][1] = $Node.text
	;Атрибуты
	Local $Prop = $Node.attributes
	For $i=0 to $Prop.length-1
		Redim $Attribs[ UBound($Attribs)+1 ][2]
		$Attribs[ UBound($Attribs)-1 ][0] = $Prop.Item($i).name
		$Attribs[ UBound($Attribs)-1 ][1] = $Prop.Item($i).text
	Next
	;Проверим дочерние узлы
	Local $child = $Node.childNodes
	For $i=0 to $child.length-1
		If $Child.item($i).tagName = "" then ContinueLoop
		Redim $Attribs[ UBound($Attribs)+1 ][2]
		$Attribs[ UBound($Attribs)-1 ][0] = $Child.item($i).tagName
		$Attribs[ UBound($Attribs)-1 ][1] = $oSelf.getData($Child.item($i))
	Next
	;Возвращаем собранный массив
	return $Attribs
EndFunc