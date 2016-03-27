#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <Debug.au3>
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
	If not FileExists($Path) then
		msgbox(32,"Информация", "Не удалось создать файл:" & $Path)
	EndIf
	return $Path
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

Func Template_TemplateEditor()
	Local $tmp = Template_getTemplateList()
	Local $Form1 = GUICreate("Редактор", 368, 248, 192, 124)
	Local $Group1 = GUICtrlCreateGroup("Список шаблонов", 8, 8, 353, 137)
	Local $List1 = GUICtrlCreateListView("Список файлов шаблонов", 16, 24, 249, 110)
	Local $Button1 = GUICtrlCreateButton("Создать", 272, 24, 75, 25)
	Local $Button2 = GUICtrlCreateButton("Выбрать", 272, 56, 75, 25)
	Local $Button3 = GUICtrlCreateButton("Удалить", 272, 88, 75, 25)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Local $Group2 = GUICtrlCreateGroup("Соответствие атрибутов шаблонов", 8, 152, 353, 89)
	Local $Label1 = GUICtrlCreateLabel("Входной атрибут", 16, 168, 91, 17)
	Local $Combo1 = GUICtrlCreateCombo("<>", 16, 184, 153, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	Local $Label2 = GUICtrlCreateLabel("Выходной атрибут", 184, 168, 99, 17)
	Local $Input1 = GUICtrlCreateInput("<>", 184, 184, 161, 21)
	Local $Button4 = GUICtrlCreateButton("Новый", 16, 208, 75, 25)
	Local $Button5 = GUICtrlCreateButton("Удалить", 96, 208, 75, 25)
	Local $Button6 = GUICtrlCreateButton("Утвердить шаблон", 184, 208, 163, 25)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	;Получим список доступных шаблонов
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
			Case -3
				Exit
			;ПОлучить список шаблонов
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
			;Cоздание шаблона
			Case $Button1
				Local $tmpName = InputBox("Создание шаблона","Укажите название шаблона","TemplateNAme")
				If $tmpName = "" then
					msgbox(64,"Информация","Шаблон не будет создан, так как не указано название")
					ContinueLoop
				EndIf
				Local $tmpFile = Template_createTemplateFile($tmpName)
				If not FileExists($tmpFile) Then
					msgbox(64,"Информация","Не удалось создать новый файл шаблона, файл не добавлен")
					ContinueLoop
				EndIf
				_GUICtrlListView_DeleteAllItems($List1)
				$tmp = Template_getTemplateList()
				If IsArray($tmp) Then
					For $i=1 to $tmp[0]
						ConsoleWrite($tmp[$i]& @CRLF)
						GUICtrlCreateListViewItem($tmp[$i],$List1)
					Next
				EndIf
			;Получить набор входных атрибутов
			Case $Button2
				GUICtrlSetData($Combo1,"")
				GUICtrlSetData($Input1,"")
				Local $src = _Source()
				Local $SelectedPath = _GUICtrlListView_GetSelectedIndices($List1)
				$SelectedPath 		= _GUICtrlListView_GetItemTextString($List1, Int($SelectedPath))
				$src.setSrc($SelectedPath)
				$src.setModel( _Source_XML() )	;Выбираем тип работы с источником: XML
				$src.open()
				Local $items = $src.getData()	;Получаем данные из источника
				_DebugOut("+> Отобразим элементы: "& UBound($items)-1)
				For $i=0 to UBound($items)-1
					Local $Item = $items[$i][1] ;Узел атрибута
					If not IsArray($Item) then ContinueLoop
					GuiCtrlSetData($Combo1,$Item[2][1])
					ConsoleWrite($Item[2][1] & @CRLF)
				Next
			;Получить значение
			Case $Combo1
				Local $Node = $src.select("Templates/Property[@name='"& GUICtrlRead($Combo1) &"']")
				If not IsArray($Node) then ContinueLoop
				GUICtrlSetData($Input1,$Node[3][1])
			;Удалить значение
			Case $Button5
				Local $Node = $src.selectNode("./Property[@name='"& GUICtrlRead($Combo1) &"']")
				$src.deleteNode()
				GUICtrlSetData($Combo1,"")
				GUICtrlSetData($Input1,"")
			;Сохранение изменений
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
	;Передаём теперь шаблон
	Local $SelectedPath = _GUICtrlListView_GetSelectedIndices($List1)
	$SelectedPath 		= _GUICtrlListView_GetItemTextString($List1, Int($SelectedPath))
	$src.setSrc($SelectedPath)
	$src.setModel( _Source_XML() )	;Выбираем тип работы с источником: XML
	$src.open()
	Local $items = $src.getData()	;Получаем данные из источника
	;Закрываем окно
	GUIDelete($Form1)
	return $src
EndFunc

;Редактор элемента шаблона
Func Main_Template_Edit()
	Local $Param, $Value
	Local $Form1 = GUICreate("Редактор единиц шаблона", 264, 119, 192, 124)
	Local $Group2 = GUICtrlCreateGroup("Шаблон", 8, 8, 249, 105)
	Local $Label1 = GUICtrlCreateLabel("Входной атрибут", 16, 32, 92, 17)
	Local $Label2 = GUICtrlCreateLabel("Выходной атрибут", 16, 64, 105, 17)
	Local $Input = GUICtrlCreateInput("MileStone", 128, 24, 120, 21)
	Local $Output = GUICtrlCreateInput("Title", 128, 56, 121, 21)
	Local $Button1 = GUICtrlCreateButton("Сохранить", 16, 80, 75, 25)
	Local $Button2 = GUICtrlCreateButton("Отменить", 176, 80, 75, 25)
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