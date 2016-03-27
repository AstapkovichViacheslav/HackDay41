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
		.AddMethod("save", 		"_SourceXML_save")
		.AddMethod("isReady",	"_SourceXML_isReady")
		.AddMethod("getData",	"_SourceXML_getData");

		.AddMethod("createNode","_SourceXML_createNode")
		.AddMethod("selectNode","_SourceXML_selectNode")
		.AddMethod("deleteNode","_SourceXML_deleteNode")
		.AddMethod("setParam",	"_SourceXML_setParam")
		.AddMethod("getParams",	"_SourceXML_getParams")

        .AddProperty("obj", $ELSCOPE_PRIVATE, ObjCreate("Microsoft.XMLDOM") )
		.AddProperty("src", $ELSCOPE_PRIVATE, "")
		.AddProperty("ptr", $ELSCOPE_PRIVATE, "")
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
Func _SourceXML_getSrc($oSelf)
	#forceref $oSelf
	return $oSelf.src
EndFunc

;Указать модель работы с данными
Func _SourceXML_create($oSelf)
	#forceref $oSelf
	IF fileExists($oSelf.src) then FileDelete($oSelf.src)
	return $oSelf.save()
EndFunc

;сохранить изменения
Func _SourceXML_save($oSelf)
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
	If not IsObj($Node) Then
		ConsoleWrite("!> Ошибка! Узел получения данных не найден!" & @CRLF)
		return 0
	Endif
	IF StringLen($Node.tagName) = 0 then
		ConsoleWrite("!> Ошибка определения узла получения данных." & @CRLF)
		return 0
	EndIf
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
		ConsoleWrite($Child.item($i).tagName & @CRLF)
	Next
	;Возвращаем собранный массив
	If $oNode="" then _SourceXML_printData($Attribs)
	return $Attribs
EndFunc

Func _SourceXML_printData($Attribs, $level=0)
	Local $cnt = UBound($Attribs)-1
	;ConsoleWrite("<<<PRINT DATA>>> Items:"& $cnt & @CRLF)
	For $I=0 to $cnt
		For $z=1 to $level
			ConsoleWrite(" ")
		Next
		If not IsArray($Attribs[$i][1]) then
			ConsoleWrite("["&$i&"]"&$Attribs[$i][0]&":"&$Attribs[$i][1] & @CRLF)
		Else
			Local $Item = $Attribs[$i][1]
			ConsoleWrite("["&$i&"]"&$Attribs[$i][0] & @CRLF)
			_SourceXML_printData($Item, $level+1)
		EndIf
	Next
	;TODO
EndFunc

Func _SourceXML_createNode($oSelf, $Node)
	#forceref $oSelf
	If not IsOBj($NOde) then
		ConsoleWrite("!> Не передан объект!")
		return 0
	EndIf
	;Определим узел, для которого будет создавать узел
	Local $Root
	;Если есть точка отсчёта - всё просто
	If $oSelf.ptr <> "" then $Root = $oSelf.ptr
	;Если точки нет
	If $oSelf.ptr = "" Then
		$Root = $oSelf.obj.documentElement
		;Если не нашли корневого элемента - корнем будет сам документ
		If not IsObj($Root) or StringLen($Root.tagName)=0 then $Root = $oSelf.obj
	EndIf
	;Проверим, что удалось указать рут
	If not IsObj($Root) Then
		_DebugOut("!> Для создания нода не удалось определить родитея")
		Return 0
	EndIf
	ConsoleWrite("+> Root is: "& $root.tagName & @CRLF)
	;Обработаем, если нам передали инстукцию
	If $Node.exists("ProcessingInstruction") Then
		Local $Split = StringSPlit($Node("ProcessingInstruction")," ")
		If not IsArray($Split) then return 0
		If UBound($Split) <> 3 Then return 0
		Local $Instr = $oSelf.obj.createProcessingInstruction( $Split[1],$Split[2])
		$oSelf.obj.insertBefore($Instr,$oSelf.obj.childNodes(1))
		_DebugOut("+> Добавлен инструкция '"& $Node("ProcessingInstruction") &"'")
		return 1
	EndIf
	;Если нам передали просто узел
	If not IsObj($oSelf.obj) Then
		_DebugOut("!> Ошибка obj")
		return 0
	EndIf
	If ObjName($NOde) = "IXMLDOMElement" Then
		Local $element = $Node
	Else
		Local $element = $oSelf.obj.createElement( $Node("tagName") )
		If $Node.exists("text") then $element.text = $Node("text")
	EndIf
	$Root.appendChild($element)
	;Если создаётся корневой узел - создаём XML инструкцию
	If $oSelf.ptr = "" then
		Local $Instr = $oSelf.obj.createProcessingInstruction("xml","version='1.0'")
		$oSelf.obj.insertBefore($Instr,$oSelf.obj.childNodes(0))
	EndIf
	;Сместим указатель на новый элемент
	$oSelf.ptr = $Element
	;Добавим атрибуты, если надо
	Local $keys = $Node.keys
	For $i=0 to UBound($keys)-1
		If $keys[$i] = "tagName" or $keys[$i] = "text" then ContinueLoop
		$element.setAttribute($keys[$i], $Node($keys[$i]) );
	Next
	return 1
EndFunc

Func _SourceXML_selectNode($oSelf, $Node)
	#forceref $oSelf
	Local $Root
	If $oSelf.ptr = "" then
		$Root = $oSelf.obj.documentElement
	Else
		$Root = $oSelf.ptr
	EndIf
	;Способ получить родителя
	If $Node="" then
		$Root = $oSelf.obj.documentElement
		$oSelf.ptr = $Root
		return $Root
	EndIf
	If $Node=".." then
		$Root = $oSelf.ptr.parentNode
		$oSelf.ptr = $Root
		return $Root
	EndIf
	If not IsObj($Root) then $Root = $oSelf.obj
	If StringLen($Root.tagName) = 0 then $Root = $oSelf.obj

	ConsoleWrite("RootNode: "& $Root.tagName & @CRLF)
	Local $Selected = $Root.selectSingleNode($Node)
	If not isObj($Selected) Then
		_DebugOut("!> Не удалось выбрать узел '"& $Node &"'")
		return 0
	EndIf
	$oSelf.ptr = $Selected
	return $Selected
EndFunc

Func _SourceXML_deleteNode($oSelf)
	#forceref $oSelf
	If $oSelf.ptr = "" then
		_DebugOut("!> Не выбран узел ($obj.selectNode)")
		return 0
	EndIf
	local $Parent = $oSelf.ptr.parentNode
	msgbox(0,"test", $Parent.tagName)
	$Parent.removeChild($oSelf.ptr)
	$oSelf.ptr = $Parent
	$oSelf.save()
	return 1
EndFunc

Func _SourceXML_setParam($oSelf, $Param, $Value)
	#forceref $oSelf
	If $oSelf.ptr = "" then
		_DebugOut("!> Не выбран узел ($obj.selectNode)")
		return 0
	EndIf
	$oSelf.ptr.setAttribute($Param, $Value)
	$oSelf.save()
	return 1
EndFunc

Func _SourceXML_getParams($oSelf)
	#forceref $oSelf
	If $oSelf.ptr = "" then
		_DebugOut("!> Не выбран узел ($obj.selectNode)")
		return 0
	EndIf
	;Получаем атрибуты
	Local $Attribs[1][2]
	Local $Prop = $oSelf.ptr.attributes
	For $i=0 to $Prop.length-1
		Redim $Attribs[ UBound($Attribs)+1 ][2]
		$Attribs[ UBound($Attribs)-1 ][0] = $Prop.Item($i).name
		$Attribs[ UBound($Attribs)-1 ][1] = $Prop.Item($i).text
	Next
	return $Attribs
EndFunc