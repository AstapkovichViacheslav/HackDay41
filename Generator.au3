#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here



Func Generator_createDoc($PtrnSrc, $Path = "C:\testD.xml")
	;Создаём документ и открываем
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc($Path)			;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.create()
	$src.open
	;Создаём основу
	Local $Node 	= $PtrnSrc.selectNode("")
	Local $RootNode = $src.element($Node.tagName)

	Local $Params 	= $PtrnSrc.getParams()
	For $i=0 to UBound($PArams)-1
		$RootNode.add($Params[$i][0], $Params[$i][1])
	Next
	$src.createNode($RootNode)
	$src.save()
	;Оформим настройки документа
	Local $DocProp = $src.element("o:DocumentProperties")
	$src.createNode($DocProp)

	Local $Node 	= $PtrnSrc.selectNode("o:DocumentProperties")
	Local $Params 	= $Node.childNodes

	For $i=1 to $Params.length-1
		$Node = $src.element($Params.Item($i).tagName)
		$src.createNode($Node)
		$src.selectNode("..")
	Next
	$src.save()
	;Скопируем из шаблона настройки стилей и оформлеие
	Local $El = ["w:fonts","w:lists","w:styles","w:shapeDefaults","w:docPr"]
	For $i=0 to UBound($El)-1
		$PtrnSrc.selectNode("")
		$Src.selectNode("")
		Local $CopyNode = $PtrnSrc.selectNode($El[$i])
		$src.createNode($CopyNode)
	Next
	$src.save()
	;Подготовим нод для тела документа
	$PtrnSrc.selectNode("")
	$Src.selectNode("")

	$Node = $src.element("w:body")	;<BODY>
	$src.createNode($Node)

	$Node = $src.element("wx:sect")	;<SRC>
	$src.createNode($Node)
	;Скопируем настройки
	Local $SectPR = $PtrnSrc.selectNode("w:body/wx:sect/w:sectPr")
	$src.createNode($SectPR)
	$src.save()
	return $src
EndFunc

Func Generator_addPatternPart($PtrnSrc, $Path)
	;Создаём документ и открываем
	Local $src = _Source()			;Создаю ресурс
	$src.setSrc($Path)				;Указываю путь
	$src.setModel( _Source_XML() )	;Указываю способ обработки
	$src.open
	$src.selectNode("")
	$PtrnSrc.selectNode("")
	Local $PatternRoot 	= $PtrnSrc.selectNode("w:body/wx:sect")
	Local $DocRoot 		= $src.selectNode("w:body/wx:sect")
	MSgbox(0,"test", $DocRoot.tagName & @TAB & $PatternRoot.tagName)
	;Создаём шаблонную часть
	;Идём по всем секциям
	Local $Items 		= $PatternRoot.childNodes
	msgbox(0,"test count", $Items.length-1)

	For $i=1 to $Items.length-1
		Local $Item = $Items.Item($i)
		If not IsObj($Item) then ContinueLoop
		msgbox(0,"test", $Item.tagName)

		If $Item.tagName = "w:sectPr" then ContinueLoop

		$src.createNode($Item)
		$src.selectNode("..")
	Next
	$src.save()
EndFunc