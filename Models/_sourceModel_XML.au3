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

;��������, ��� ��� ������ XML
Func _SourceXML_getType($oSelf)
	#forceref $oSelf
	return "_sourceModel"
EndFunc

;������� ���� � ���������
Func _SourceXML_setSrc($oSelf, $Path)
	#forceref $oSelf
	$oSelf.src = $Path
EndFunc

;�������� ���� � ���������
Func _SourceXML_getSrc($oSelf, $Path)
	#forceref $oSelf
	return $oSelf.src
EndFunc

;������� ������ ������ � �������
Func _SourceXML_create($oSelf)
	#forceref $oSelf
	Local $code = $oSelf.obj.save($oSelf.src)
	return FileExists($oSelf.src)
EndFunc

;����� �� �������� ������
Func _SourceXML_isReady($oSelf)
	return FileExists($oSelf.src)
EndFunc

;�������� ���������
Func _SourceXML_open($oSelf)
	#forceref $oSelf
	if $oSelf.src = "" then
		ConsoleWrite("!> XML: �� ������ ����" & @CRLF)
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
	;����������� �������� ����
	$Attribs[0][0] = "tagName"
	$Attribs[0][1] = $Node.tagName
	$Attribs[1][0] = "text"
	Local $RegExp = StringRegExp($Node.xml,"<\w",3)
	if UBound($RegExp)-1 = 0 then $Attribs[1][1] = $Node.text
	;��������
	Local $Prop = $Node.attributes
	For $i=0 to $Prop.length-1
		Redim $Attribs[ UBound($Attribs)+1 ][2]
		$Attribs[ UBound($Attribs)-1 ][0] = $Prop.Item($i).name
		$Attribs[ UBound($Attribs)-1 ][1] = $Prop.Item($i).text
	Next
	;�������� �������� ����
	Local $child = $Node.childNodes
	For $i=0 to $child.length-1
		If $Child.item($i).tagName = "" then ContinueLoop
		Redim $Attribs[ UBound($Attribs)+1 ][2]
		$Attribs[ UBound($Attribs)-1 ][0] = $Child.item($i).tagName
		$Attribs[ UBound($Attribs)-1 ][1] = $oSelf.getData($Child.item($i))
	Next
	;���������� ��������� ������
	return $Attribs
EndFunc