#cs


#ce





Func _Source($SourcePath="")
    Local $oClass = _AutoItObject_Class()
    With $oClass
		.AddMethod("setSrc", 	"_Source_setSrc")		;Указать путь к источнику
		.AddMethod("getSrc", 	"_Source_getSrc")		;Получить путь к источнику
        .AddMethod("setModel", 	"_Source_setModel")		;Указать библиотеку модели (XML, TXT и т.п.)
        .AddMethod("create", 	"_Source_create")		;Создать источник
		.AddMethod("open", 		"_Source_open")			;Открыть источник
		.AddMethod("save", 		"_Source_save")			;Сохранить изменеия в источнике
		.AddMethod("isReady",	"_Source_isReady")		;Готов ли источник к работе
		.AddMethod("getData",	"_Source_getData")		;Получить данные источника
		.AddMethod("select",	"_Source_selectData")	;Выбрать в данных источника узел
		.AddMethod("element",	"_Source_createElement");Создаь элемен


		.AddMethod("createNode","_Source_createNode")	;Создание узла
		.AddMethod("selectNode","_Source_selectNode")	;Выбор узла (переход в узел)
		.AddMethod("setParam",	"_Source_setParam")		;Указать атрибуты/параметры узла

		.AddProperty("src", 	$ELSCOPE_PRIVATE, $SourcePath)
        .AddProperty("model", 	$ELSCOPE_PRIVATE, null)
    EndWith
	_DebugOut("+> Создан объект Source")
	return $oClass.Object
EndFunc

;https://msdn.microsoft.com/en-us/library/x4k5wbx4(v=vs.84).aspx
Func _Source_createElement($oSelf, $tagName, $tag = "tagName")
	#forceref $oSelf
	Local $Element = ObjCreate("Scripting.Dictionary")
	$Element.add($tag, $tagName)
	return $Element
EndFunc

;Указать путь к источнику
Func _Source_setSrc($oSelf, $Path)
	#forceref $oSelf
	$oSelf.src = $Path
	_DebugOut("+> Путь к источнику: "& $oSelf.src)
EndFunc

;Получить путь к источнику
Func _Source_getSrc($oSelf, $Path)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	return $oSelf.model.getSrc()
EndFunc

;Указать модель работы с данными
Func _Source_setModel($oSelf, $ModelObj)
	#forceref $oSelf
	If not IsObj($ModelObj) Then
		_DebugOut("!> Для модели не передан объект")
		Return
	EndIf
	If $ModelObj.getType <> "_sourceModel" Then
		_DebugOut("!> Передан некорректный объект")
		Return
	EndIf
	$oSelf.model = $ModelObj
	if $oSelf.src <> "" then $oSelf.model.setSrc($oSelf.src)
	_DebugOut("+> Выбрана модель для работы с источником")
EndFunc

;Создать источник для работы с данными
Func _Source_create($oSelf)
	#forceref $oSelf
	If $oSelf.src = "" Then
		_DebugOut("!> Для создания не указан путь!")
		Return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $Ret = $oSelf.model.create()
	_DebugOut("+> Создан источник: "&$Ret)
	return $Ret
EndFunc

;Сохранить изменения в источнике
Func _Source_save($oSelf)
	#forceref $oSelf
	If $oSelf.src = "" Then
		_DebugOut("!> Для сохранения не указан путь!")
		Return
	EndIf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $Ret = $oSelf.model.save()
	_DebugOut("+> Изменения данных сохранены: "&$Ret)
	return $Ret
EndFunc

;Открыть источник для работы с данными
Func _Source_open($oSelf)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> Источник данных не готов!")
		Return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $code = $oSelf.model.open()
	_DebugOut("+> Источник данных открыт:"&$code)
EndFunc

Func _Source_getData($oSelf)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> Источник данных не готов!")
		Return
	EndIf
	Local $Data = $oSelf.model.getData()
	_DebugOut("+> Получены данные:"& IsArray($Data) )
	return $Data
EndFunc

Func _Source_selectData($oSelf, $XPath, $dataArray="")
	#forceref $oSelf
	Local $Data
	If IsArray($dataArray) Then
		$Data = $dataArray
	Else
		$Data = $oSelf.getData()
	EndIf
	IF $XPAth = "" then return $Data
	Local $cTag
	Local $Path = StringSplit($XPath,"/")
	_DebugOut("+> Поиск элемента: '"& $XPath &"'")
	For $i=1 to $Path[0]
		;Опредеим искомый нод и его атрибуты
		Local $params = StringRegExp($Path[$i],"(?U)(@.*='.*')",3)
		If IsArray($params) Then
			$cTag = StringMid($Path[$i],1, StringInStr($Path[$i],"[")-1)
		Else
			$cTag = $Path[$i]
		EndIf
		;ConsoleWrite($cTag & @TAB & UBound($params) & @TAB & IsArray($params) & @CRLF)
		;Если это корневой элемент - идём дальше
		If $i=1 and $cTag = $Data[0][1] and not IsArray($params) then
			ContinueLoop
		EndIf
		;Ищем среди дочерних нодов указанный
		For $z=2 to UBound($Data)-1
			If $Data[$z][0] <> $cTag then ContinueLoop

			Local $Item = $Data[$z][1]
			If not IsArray($Params) then
				$Data = $Item ;Параметры не указны - берём что нашли
				ContinueLoop 2
			EndIf
			;Если указаны атрибуты - они должны все совпасть для узла
			For $p=0 to UBound($params)-1
				Local $split=StringSPlit($params[$p],"=") ;[папаметр, значение]
				$split[1] = StringMid($split[1],2, StringLen($split[1]))
				$split[2] = StringMid($split[2],2, StringLen($split[2])-2)
				;ConsoleWrite("Ищу параметр '"&$split[1] & " со значением " & $split[2] & " среди элементов: "& UBound($Item)-1& @CRLF)
				;Ищем среди атрибутов узла текущий атрибут
				Local $found = False
				For $a=2 to UBound($Item)-1
					If StringLen($Item[$a][0]) = 0 then ContinueLoop
					IF $Item[$a][0] = $split[1] and $Item[$a][1]= $split[2] Then
						;ConsoleWrite("Нашли" & @CRLF)
						$found = True
						ExitLoop
					EndIf
				Next
				If not $found then ContinueLoop 2 ;Если не нашли - переходим к обработке следующего дочернего нода
			Next
			;Узел найден, атрибуты все совпадают
			$Data = $Item
			ContinueLoop 2
		Next
		;Если дошло до сюда - путь не найден
		_DebugOut("!> Не удалось найти элемент: '"& $cTag &"'!")
		Return
	Next
	return $Data
EndFunc

Func _Source_createNode($oSelf, $Node)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	If not IsObj($Node) Then
		_DebugOut("!> Для работы с данными не передан src.element!")
		return
	EndIf
	If ObjName($Node) <> "Dictionary" Then
		_DebugOut("!> Для работы с данными передан некорректный объект!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> Источник данных не готов!")
		Return
	EndIf
	Local $code = $oSelf.model.createNode($Node)
	_DebugOut("+> Создание узла:"& IsArray($code) )
	return $code
EndFunc

Func _Source_selectNode($oSelf, $Path)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	Local $code = $oSelf.model.selectNode($Path)
	_DebugOut("+> Выбор узла '"& $Path &"':"& IsObj($code) )
	return $code
EndFunc

Func _Source_setParam($oSelf, $Param, $Value)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> Для работы с данными не указана модель!")
		return
	EndIf
	Local $code = $oSelf.model.setParam($Param, $Value)
	_DebugOut("+> Изменение параметра '"& $Param &"'" )
	return $code
EndFunc


Func WriteTest()
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc("C:\test.xml")		;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.create()
	$src.open()
	Local $Node = $src.element("test1")
	$Node.add("param1","pval1")
	$Node.add("param2","pval2")
	$src.createNode($Node)
	Local $Node = $src.element("test2")
	$src.createNode($Node)
	$src.selectNode("..")
	Local $Node = $src.element("test3")
	$src.createNode($Node)
	$src.setParam("param2","pval2-2")

	Local $Instr =$src.element("test tst='f'","ProcessingInstruction")
	$src.createNode($Instr)




	$src.save()
EndFunc

Func ReadTest()
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc("C:\pim.xml")		;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.open()						;Открываю источник
	Local $PIM = $src.getData()		;Получаю данные
	;Ищу строку без параметров
	Local $String = "w:wordDocument/o:DocumentProperties/o:Title"
	Local $Node = $src.select($String,$PIM)
	ConsoleWrite("1 test: "& $Node[1][1] & @CRLF)
	;Ищу строку с параметром
	Local $String = "w:wordDocument/w:styles/w:style[@w:styleId='2']/wx:uiName"
	Local $Node = $src.select($String,$PIM)
	ConsoleWrite("2 test: "& $Node[2][1] & @CRLF)
EndFunc
