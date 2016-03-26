#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Isakov, Astapkovich, Sejfulaev

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#include <Array.au3>
#include <FileConstants.au3>
#include <StringConstants.au3>

#region GUI

Global $InstallForm ; ���� ������������
Global $Logo 		; �������
Global $NextButton  ; ������ �����
Global $Text		; �������� ����� ���� (���������� �������� ���� ���������� ������ ��������)

; ---------- �������� GUI ��� ����������� ������ ������������ ---------
Global $Steps_Label ; ����� � ������������ ���� ���������

Global $Org_Label
Global $F_Label
Global $I_Label
Global $O_Label

Global $Org_Input
Global $F_Input
Global $I_Input
Global $O_Input
; ---------------------------------------------------------------------

; ������ �� ��������� ������ ������ AutoPIM
Global $arrRegDelete[0]
Global $arrDirDelete[0]
Global $arrFIO

#EndRegion

; ������� ����������� ������ ������������
Func SetUserInfo($bShow)
	local $arr_fio[3]

	If $bShow Then
		$Org_Label = GUICtrlCreateLabel("�����������", 43, 122, 71, 17)
		$Steps_Label = GUICtrlCreateLabel("", 24, 96, 449, 17)
		$F_Label = GUICtrlCreateLabel("�������", 60, 153, 53, 17)
		$I_Label = GUICtrlCreateLabel("���", 85, 183, 26, 17)
		$O_Label = GUICtrlCreateLabel("��������", 62, 213, 51, 17)
		$Org_Input = GUICtrlCreateInput("", 120, 120, 250, 21)


		local $count = Ubound($arrFIO)
		if $count > 3 Then $count = 3

		For $i = 0 To $count - 1  Step 1
			$arr_fio[$i] = $arrFIO[$i]
		Next

		$F_Input = GUICtrlCreateInput($arr_fio[0], 120, 150, 121, 21)
		$I_Input = GUICtrlCreateInput($arr_fio[1], 120, 180, 121, 21)
		$O_Input = GUICtrlCreateInput($arr_fio[2], 120, 210, 121, 21)
	Else

		GUICtrlSetState($Org_Label, $GUI_HIDE)
		GUICtrlSetState($F_Label, $GUI_HIDE)
		GUICtrlSetState($I_Label, $GUI_HIDE)
		GUICtrlSetState($O_Label, $GUI_HIDE)
		GUICtrlSetState($Org_Input, $GUI_HIDE)
		GUICtrlSetState($F_Input, $GUI_HIDE)
		GUICtrlSetState($I_Input, $GUI_HIDE)
		GUICtrlSetState($O_Input, $GUI_HIDE)

	EndIf
EndFunc

Func Install()

	Local $installStep = 1;
	$InstallForm = GUICreate("���������  �������", 501, 279, 413, 190)
	$Logo = GUICtrlCreatePic(".\Images\�������.jpg", 0, 0, 500, 85)
	$NextButton = GUICtrlCreateButton("�����", 368, 248, 123, 25)

	$Text = GUICtrlCreateEdit("", 24, 96, 449, 129, BitOR($ES_READONLY,$ES_WANTRETURN), 0)
	GUICtrlSetData(-1, StringFormat("��� ������������ ����������� ��������� �������\r\n\r\n������� �������� ������� ������ ������������ �� ��������� ������ ��������.\r\n��������� ����� ����������� ������� � ������� ��� �������� ������������.\r\n\r\n� �������� ��������� ��������� ������� �������� ����� ������������� ������. \r\n��� ������� ������ ������������� ������ �������� ����������. \r\n\r\n��� �����������  ��������� ������� ������ "&Chr(34)&"�����"&Chr(34)&""))

	GUISetState(@SW_SHOW)

	While 1
		local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $NextButton
				Switch $installStep
					Case 1
						$installStep = $installStep + 1
						GUICtrlSetData($Text, StringFormat("��� 1: ����� ����� ������������� ������ �������" & findPreviouslyInstalled()))

					Case 2
						$installStep = $installStep + 1
						GUICtrlSetState($Text, $GUI_HIDE)
						SetUserInfo(True)
						GUICtrlSetData($Steps_Label, StringFormat("��� 2: ����������� ������ � ������������ � �������"))
					Case 3
						SetUserInfo(False)
						GUICtrlSetData($Steps_Label, StringFormat("��� 3: ���������� ���������"))
				EndSwitch
		EndSwitch
	WEnd
EndFunc

; ������� ������ ���������� ������
Func findPreviouslyInstalled()
	local $message

	; ������ ����� ������������� ������ AutoPIM
	Local $arrType[3] = ["*","Directory", ".xml"]
	Local $foundPath[0]
	Local $Key

	For $Type in $arrType
		$Key = "HKEY_CLASSES_ROOT\"&$Type&"\Shell\AutoPIM"
		If RegRead($Key, "MUIVerb") <> "" Then
			_ArrayAdd($arrRegDelete, $Key)
			local $path = RegRead($Key&"\command", "")
			; ���� �� ����� ���� �� �������� ���
			if $path <> "" Then
				$path = StringLeft($path, StringInStr($path, "AutoPIM.exe") - 2)
				if FileExists($path) Then
					if _ArraySearch($arrDirDelete, $path) = -1 Then
						_ArrayAdd($arrDirDelete, $path)
					EndIf
				EndIf
			EndIf
		EndIf
	Next

	If Ubound($arrRegDelete) or Ubound($arrDirDelete) Then
		$message = "\r\n\r\n������� ����� ������������� ������ AutoPIM:\r\n"
		For $dir in $arrDirDelete
			$message = $message & "\r\n���� ���������: " & $dir
			; ��������� ��������� �� ����� opt.ini ��� ������ ������ AutoPIM
			Local $hFileOpen = FileOpen($dir & "\Opt.ini", $FO_READ)
			If $hFileOpen <> -1 Then
				Local 	$str = StringReplace(FileReadLine($hFileOpen,2), "FIO=", "")
						$str = StringStripWS($str, $STR_STRIPSPACES)
				$arrFIO = StringSplit($str, " ", $STR_NOCOUNT)
				FileClose ($hFileOpen)
			EndIf
		Next

		For $reg in $arrRegDelete
			$message = $message & "\r\n���� � �������: " & $reg
		Next

		$message = $message & "\r\n\r\n�����: � ���� ��������� ��������� ������ � ������� ����� �������"
	Else
		$message = "\r\n����� ������������� ������ AutoPIM �� �������"
	EndIf

	return $message
EndFunc

Func CreateMenuItem($GroupName)
	Local $Type = "xml"
	Local $Key = "HKEY_CLASSES_ROOT\."&$Type&"\Shell\"&@ScriptName
	RegWrite($key)
	RegWrite($key,"MUIVerb","REG_SZ",$GroupName)
	RegWrite($key,"SubCommands","REG_SZ",$GroupName&"1" &";"& $Groupname&"2")
	;http://rapidsoft.org/articles/wintuning/item/101-context_menu_section
EndFunc