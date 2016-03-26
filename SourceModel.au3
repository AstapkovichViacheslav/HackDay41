#cs


#ce





Func _Source($SourcePath="")
    Local $oClass = _AutoItObject_Class()
    With $oClass
		.AddMethod("setSrc", 	"_Source_setSrc")		;������� ���� � ���������
		.AddMethod("getSrc", 	"_Source_getSrc")		;�������� ���� � ���������
        .AddMethod("setModel", 	"_Source_setModel")		;������� ���������� ������ (XML, TXT � �.�.)
        .AddMethod("create", 	"_Source_create")		;������� ��������
		.AddMethod("open", 		"_Source_open")			;������� ��������
		.AddMethod("save", 		"_Source_save")			;��������� �������� � ���������
		.AddMethod("isReady",	"_Source_isReady")		;����� �� �������� � ������
		.AddMethod("getData",	"_Source_getData")		;�������� ������ ���������
		.AddMethod("select",	"_Source_selectData")	;������� � ������ ��������� ����
		.AddMethod("element",	"_Source_createElement");������ ������


		.AddMethod("createNode","_Source_createNode")	;�������� ����
		.AddMethod("selectNode","_Source_selectNode")	;����� ���� (������� � ����)
		.AddMethod("setParam",	"_Source_setParam")		;������� ��������/��������� ����

		.AddProperty("src", 	$ELSCOPE_PRIVATE, $SourcePath)
        .AddProperty("model", 	$ELSCOPE_PRIVATE, null)
    EndWith
	_DebugOut("+> ������ ������ Source")
	return $oClass.Object
EndFunc

;https://msdn.microsoft.com/en-us/library/x4k5wbx4(v=vs.84).aspx
Func _Source_createElement($oSelf, $tagName, $tag = "tagName")
	#forceref $oSelf
	Local $Element = ObjCreate("Scripting.Dictionary")
	$Element.add($tag, $tagName)
	return $Element
EndFunc

;������� ���� � ���������
Func _Source_setSrc($oSelf, $Path)
	#forceref $oSelf
	$oSelf.src = $Path
	_DebugOut("+> ���� � ���������: "& $oSelf.src)
EndFunc

;�������� ���� � ���������
Func _Source_getSrc($oSelf, $Path)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	return $oSelf.model.getSrc()
EndFunc

;������� ������ ������ � �������
Func _Source_setModel($oSelf, $ModelObj)
	#forceref $oSelf
	If not IsObj($ModelObj) Then
		_DebugOut("!> ��� ������ �� ������� ������")
		Return
	EndIf
	If $ModelObj.getType <> "_sourceModel" Then
		_DebugOut("!> ������� ������������ ������")
		Return
	EndIf
	$oSelf.model = $ModelObj
	if $oSelf.src <> "" then $oSelf.model.setSrc($oSelf.src)
	_DebugOut("+> ������� ������ ��� ������ � ����������")
EndFunc

;������� �������� ��� ������ � �������
Func _Source_create($oSelf)
	#forceref $oSelf
	If $oSelf.src = "" Then
		_DebugOut("!> ��� �������� �� ������ ����!")
		Return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $Ret = $oSelf.model.create()
	_DebugOut("+> ������ ��������: "&$Ret)
	return $Ret
EndFunc

;��������� ��������� � ���������
Func _Source_save($oSelf)
	#forceref $oSelf
	If $oSelf.src = "" Then
		_DebugOut("!> ��� ���������� �� ������ ����!")
		Return
	EndIf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $Ret = $oSelf.model.save()
	_DebugOut("+> ��������� ������ ���������: "&$Ret)
	return $Ret
EndFunc

;������� �������� ��� ������ � �������
Func _Source_open($oSelf)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> �������� ������ �� �����!")
		Return
	EndIf
	If $oSelf.model.getSrc() = "" Then
		$oSelf.model.setSrc($oSelf.src)
	EndIf
	Local $code = $oSelf.model.open()
	_DebugOut("+> �������� ������ ������:"&$code)
EndFunc

Func _Source_getData($oSelf)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> �������� ������ �� �����!")
		Return
	EndIf
	Local $Data = $oSelf.model.getData()
	_DebugOut("+> �������� ������:"& IsArray($Data) )
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
	_DebugOut("+> ����� ��������: '"& $XPath &"'")
	For $i=1 to $Path[0]
		;�������� ������� ��� � ��� ��������
		Local $params = StringRegExp($Path[$i],"(?U)(@.*='.*')",3)
		If IsArray($params) Then
			$cTag = StringMid($Path[$i],1, StringInStr($Path[$i],"[")-1)
		Else
			$cTag = $Path[$i]
		EndIf
		;ConsoleWrite($cTag & @TAB & UBound($params) & @TAB & IsArray($params) & @CRLF)
		;���� ��� �������� ������� - ��� ������
		If $i=1 and $cTag = $Data[0][1] and not IsArray($params) then
			ContinueLoop
		EndIf
		;���� ����� �������� ����� ���������
		For $z=2 to UBound($Data)-1
			If $Data[$z][0] <> $cTag then ContinueLoop

			Local $Item = $Data[$z][1]
			If not IsArray($Params) then
				$Data = $Item ;��������� �� ������ - ���� ��� �����
				ContinueLoop 2
			EndIf
			;���� ������� �������� - ��� ������ ��� �������� ��� ����
			For $p=0 to UBound($params)-1
				Local $split=StringSPlit($params[$p],"=") ;[��������, ��������]
				$split[1] = StringMid($split[1],2, StringLen($split[1]))
				$split[2] = StringMid($split[2],2, StringLen($split[2])-2)
				;ConsoleWrite("��� �������� '"&$split[1] & " �� ��������� " & $split[2] & " ����� ���������: "& UBound($Item)-1& @CRLF)
				;���� ����� ��������� ���� ������� �������
				Local $found = False
				For $a=2 to UBound($Item)-1
					If StringLen($Item[$a][0]) = 0 then ContinueLoop
					IF $Item[$a][0] = $split[1] and $Item[$a][1]= $split[2] Then
						;ConsoleWrite("�����" & @CRLF)
						$found = True
						ExitLoop
					EndIf
				Next
				If not $found then ContinueLoop 2 ;���� �� ����� - ��������� � ��������� ���������� ��������� ����
			Next
			;���� ������, �������� ��� ���������
			$Data = $Item
			ContinueLoop 2
		Next
		;���� ����� �� ���� - ���� �� ������
		_DebugOut("!> �� ������� ����� �������: '"& $cTag &"'!")
		Return
	Next
	return $Data
EndFunc

Func _Source_createNode($oSelf, $Node)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	If not IsObj($Node) Then
		_DebugOut("!> ��� ������ � ������� �� ������� src.element!")
		return
	EndIf
	If ObjName($Node) <> "Dictionary" Then
		_DebugOut("!> ��� ������ � ������� ������� ������������ ������!")
		return
	EndIf
	If not $oSelf.model.isReady Then
		_DebugOut("!> �������� ������ �� �����!")
		Return
	EndIf
	Local $code = $oSelf.model.createNode($Node)
	_DebugOut("+> �������� ����:"& IsArray($code) )
	return $code
EndFunc

Func _Source_selectNode($oSelf, $Path)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	Local $code = $oSelf.model.selectNode($Path)
	_DebugOut("+> ����� ���� '"& $Path &"':"& IsObj($code) )
	return $code
EndFunc

Func _Source_setParam($oSelf, $Param, $Value)
	#forceref $oSelf
	If not IsObj($oSelf.model) Then
		_DebugOut("!> ��� ������ � ������� �� ������� ������!")
		return
	EndIf
	Local $code = $oSelf.model.setParam($Param, $Value)
	_DebugOut("+> ��������� ��������� '"& $Param &"'" )
	return $code
EndFunc


Func WriteTest()
	Local $src = _Source()			;������ ������
	$src.setSrc("C:\test.xml")		;�������� ����
	$src.setModel( _Source_XML() )	;�������� ������ ���������
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
	Local $src = _Source()			;������ ������
	$src.setSrc("C:\pim.xml")		;�������� ����
	$src.setModel( _Source_XML() )	;�������� ������ ���������
	$src.open()						;�������� ��������
	Local $PIM = $src.getData()		;������� ������
	;��� ������ ��� ����������
	Local $String = "w:wordDocument/o:DocumentProperties/o:Title"
	Local $Node = $src.select($String,$PIM)
	ConsoleWrite("1 test: "& $Node[1][1] & @CRLF)
	;��� ������ � ����������
	Local $String = "w:wordDocument/w:styles/w:style[@w:styleId='2']/wx:uiName"
	Local $Node = $src.select($String,$PIM)
	ConsoleWrite("2 test: "& $Node[2][1] & @CRLF)
EndFunc
