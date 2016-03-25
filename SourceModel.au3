#cs


#ce
#RequireAdmin
#include-once
#include <Debug.au3>
#include <AutoItObject.au3>
#include ".\Models\_sourceModel_XML.au3"
_DebugSetup("DEBUG",true,2)
Opt("MustDeclareVars", 1)

Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")
_AutoItObject_StartUp()

Func _ErrFunc()
	ConsoleWrite("!> ERR = "& $oError.description & @CRLF)
EndFunc

Func _Source($SourcePath="")
    Local $oClass = _AutoItObject_Class()
    With $oClass
		.AddMethod("setSrc", 	"_Source_setSrc")
		.AddMethod("getSrc", 	"_Source_getSrc")
        .AddMethod("setModel", 	"_Source_setModel")
        .AddMethod("create", 	"_Source_create")
		.AddMethod("open", 		"_Source_open")
		.AddMethod("isReady",	"_Source_isReady")
		.AddMethod("createNode","_Source_createNode")
		.AddMethod("selectNode","_Source_selectNode")
		.AddMethod("getData",	"_Source_getData")
		.AddProperty("src", 	$ELSCOPE_PRIVATE, $SourcePath)
        .AddProperty("model", 	$ELSCOPE_PRIVATE, null)
    EndWith
	_DebugOut("+> Создан объект Source")
	return $oClass.Object
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

;Открыть источник для работы с данными
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
	$oSelf.model.open()
EndFunc


Local $src = _Source()
$src.setSrc("C:\test.xml")
$src.setModel( _Source_XML() )
$src.create()
$src.open()
