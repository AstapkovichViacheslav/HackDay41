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

	Local $Instr = $src.element("mso-application progid='Word.Document'","ProcessingInstruction")
	$src.createNode($Instr)
	$src.save()

	;Оформим настройки документа
	Local $DocProp = $src.element("o:DocumentProperties")
	$src.createNode($DocProp)

	Local $Node 	= $PtrnSrc.selectNode("o:DocumentProperties")
	Local $Params 	= $Node.childNodes

	For $i=1 to $Params.length-1
		If not IsObj($Params.Item($i)) then ContinueLoop
		ConsoleWrite("Add param:" & $Params.Item($i).tagName & @CRLF)
		ConsoleWrite("Add value:" & $Params.Item($i).text & @CRLF)
		$src.createNode($Params.Item($i))
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
	;Создаём шаблонную часть
	;Идём по всем секциям
	Local $Items 		= $PatternRoot.childNodes
	_DEbugOut("По шаблону будет создано узлов: " & $Items.length-1)
	For $i=0 to $Items.length-1
		Local $Item = $Items.Item($i)
		If not IsObj($Item) then ContinueLoop
		;msgbox(0,"test", $Item.tagName)

		If $Item.tagName = "w:sectPr" then ContinueLoop

		$src.createNode($Item)
		$src.selectNode("..")
	Next
	$src.save()
EndFunc


Func GenerateExcelDoc()

	; ��������� ��������� XML
	local $STYLE_TEMPLATE = "<Styles>" & _
					"<Style ss:Name=""Normal"" ss:ID=""Default""/>" & _
					"<Style ss:ID=""s70"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s71"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s72"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s73"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s74"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s75"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s76"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s77"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s78"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s79"">" & _
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
					"</Style>" & _
					"<Style ss:ID=""s80"">" & _ ; ���������� ������ ��� RMS
						"<Borders>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Bottom""/>" & _
							"<Border ss:Weight=""1"" ss:LineStyle=""Continuous"" ss:Position=""Top""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Left""/>" & _
							"<Border ss:Weight=""2"" ss:LineStyle=""Continuous"" ss:Position=""Right""/>" & _
						"</Borders>" & _
						"<Font ss:Bold=""1""/>" & _
						"<Interior ss:Color=""#D8E4BC"" ss:Pattern=""Solid""/>" & _
					"</Style>" & _				 ; ����� ���������� ������ ��� RMS
					"</Styles>"


	local $ROW_TEMPLATE = "<Row ss:Index=""ROW_INDEX"">" & _
					"<Cell ss:Index=""1"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""2"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""3"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""4"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""5"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""6"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""7"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""8"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""9"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""10"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""11"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""12"" ss:StyleID=""s77""><Data ss:Type=""String"">STEP_NUMBER</Data></Cell>" & _
					"<Cell ss:Index=""13"" ss:StyleID=""s77""><Data ss:Type=""String"">STEP_DESCRIPTION</Data></Cell>" & _
					"<Cell ss:Index=""14"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"<Cell ss:Index=""15"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
					"</Row>"


	local $HEADER_ROW_TEMPLATE = 	"<Row ss:Index=""1""/>" & _
							"<Row ss:Index=""2"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">�������� ���������</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""3""/>" & _
							"<Row ss:Index=""4"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">�������:</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String"">____________</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""5""/>" & _
							"<Row ss:Index=""6"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">������:</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String"">____________</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""7""/>" & _
							"<Row ss:Index=""8"">" & _
								"<Cell ss:StyleID=""s70"" ss:MergeAcross=""4""><Data ss:Type=""String"">���������� �� �������������� ����������</Data></Cell>" & _
								"<Cell ss:StyleID=""s70"" ss:MergeAcross=""2""><Data ss:Type=""String"">���������� �� �������� �����</Data></Cell>" & _
								"<Cell ss:StyleID=""s70"" ss:MergeAcross=""1""><Data ss:Type=""String"">����/����� ����������</Data></Cell>" & _
								"<Cell ss:StyleID=""s70"" ss:MergeAcross=""4""><Data ss:Type=""String"">����������� ����������</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""9"">" & _
								"<Cell ss:Index=""1"" ss:StyleID=""s70""><Data ss:Type=""String"">������</Data></Cell>" & _
								"<Cell ss:Index=""2"" ss:StyleID=""s70""><Data ss:Type=""String"">���</Data></Cell>" & _
								"<Cell ss:Index=""3"" ss:StyleID=""s70""><Data ss:Type=""String"">���������</Data></Cell>" & _
								"<Cell ss:Index=""4"" ss:StyleID=""s70""><Data ss:Type=""String"">�������</Data></Cell>" & _
								"<Cell ss:Index=""5"" ss:StyleID=""s70""><Data ss:Type=""String"">��.�����</Data></Cell>" & _
								"<Cell ss:Index=""6"" ss:StyleID=""s70""><Data ss:Type=""String"">��</Data></Cell>" & _
								"<Cell ss:Index=""7"" ss:StyleID=""s70""><Data ss:Type=""String"">���������</Data></Cell>" & _
								"<Cell ss:Index=""8"" ss:StyleID=""s70""><Data ss:Type=""String"">���</Data></Cell>" & _
								"<Cell ss:Index=""9"" ss:StyleID=""s70""><Data ss:Type=""String"">������</Data></Cell>" & _
								"<Cell ss:Index=""10"" ss:StyleID=""s70""><Data ss:Type=""String"">����������</Data></Cell>" & _
								"<Cell ss:Index=""11"" ss:StyleID=""s70""><Data ss:Type=""String"">������</Data></Cell>" & _
								"<Cell ss:Index=""12"" ss:StyleID=""s70""><Data ss:Type=""String"">� �� ���</Data></Cell>" & _
								"<Cell ss:Index=""13"" ss:StyleID=""s70""><Data ss:Type=""String"">��� �� ���</Data></Cell>" & _
								"<Cell ss:Index=""14"" ss:StyleID=""s70""><Data ss:Type=""String"">���������</Data></Cell>" & _
								"<Cell ss:Index=""15"" ss:StyleID=""s70""><Data ss:Type=""String"">�����������</Data></Cell>" & _
							"</Row>"

	local $FOOTER_TEMPLATE ="<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1"" ss:StyleID=""s71""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""3"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""5"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""7"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""8"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""9"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""10"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""11"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""12"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""13"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""14"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""15"" ss:StyleID=""s77""><Data ss:Type=""String""/></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX""/>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">__________</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">(���������)</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">(���)</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">(�������)</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">__________</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">(���������)</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">(���)</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">(�������)</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">__________________</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">__________</Data></Cell>" & _
							"</Row>" & _
							"<Row ss:Index=""ROW_INDEX"">" & _
								"<Cell ss:Index=""1""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""2""><Data ss:Type=""String"">(���������)</Data></Cell>" & _
								"<Cell ss:Index=""3""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""4""><Data ss:Type=""String"">(���)</Data></Cell>" & _
								"<Cell ss:Index=""5""><Data ss:Type=""String""/></Cell>" & _
								"<Cell ss:Index=""6""><Data ss:Type=""String"">(�������)</Data></Cell>" & _
							"</Row>"


	local $DOC_TEMPLATE = "<?xml version=""1.0""?><?mso-application progid='Excel.Sheet'?>" & _
						"<Workbook xmlns:html=""http://www.w3.org/TR/REC-html40"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns=""urn:schemas-microsoft-com:office:spreadsheet"">" & _
						"<DocumentProperties xmlns=""urn:schemas-microsoft-com:office:office"">" & _
							"<Author>AUTHOR_PARAM</Author>" & _
						"</DocumentProperties>" & _
						"<OfficeDocumentSettings xmlns=""urn:schemas-microsoft-com:office:office"">" & _
							"<AllowPNG/>" & _
						"</OfficeDocumentSettings>" & _
						"<ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel"">" & _
								"<WindowHeight>10320</WindowHeight>" & _
								"<WindowWidth>20670</WindowWidth>" & _
								"<WindowTopX>930</WindowTopX>" & _
								"<WindowTopY>0</WindowTopY>" & _
								"<ProtectStructure>False</ProtectStructure>" & _
								"<ProtectWindows>False</ProtectWindows>" & _
							"</ExcelWorkbook>" & _
						$STYLE_TEMPLATE & _
						"<Worksheet ss:Name=""��������"">" & _
							"<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">" & _
									"<Selected/>" & _
									"<ProtectObjects>False</ProtectObjects>" & _
									"<ProtectScenarios>False</ProtectScenarios>" & _
							"</WorksheetOptions>" & _
								"<Table x:FullRows=""1"" x:FullColumns=""1"" ss:ExpandedRowCount=""10000"" ss:ExpandedColumnCount=""15"">" & _
									"<Column ss:Width=""70"" ss:Index=""3""/>" & _
									"<Column ss:Width=""60"" ss:Index=""7""/>" & _
									"<Column ss:Width=""65"" ss:Index=""10""/>" & _
									"<Column ss:Width=""54"" ss:Index=""12""/>" & _
									"<Column ss:Width=""460"" ss:AutoFitWidth=""0""/>" & _
									"<Column ss:Width=""55""/>" & _
									"<Column ss:Width=""255"" ss:AutoFitWidth=""0""/>" & _
									$HEADER_ROW_TEMPLATE & _
									"ROW_PARAM" & _
								"</Table>" & _
							"</Worksheet>" & _
						"</Workbook>"


	; ���������� ����� �����
	Dim $temp_string, $doc_text, $footer
	$doc_text = $DOC_TEMPLATE
	$footer = $FOOTER_TEMPLATE

EndFunc