#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=PPConverter.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <GUIConstantsEx.au3>
#Include <PowerPoint.au3>
#include <Constants.au3>
#include <String.au3>
#include <Array.au3>
#include <File.au3>

Global $listing
Global $Total_Files
Global $File_Has_Been_Selected = 0
GLOBAL $DIR_MESSAGE = ""

_Menu()
  Exit

Func _Menu()
	Local $idFileMenu, $idFileItem, $idRecentFilesMenu, $idSeparator1
	Local $idExitItem, $idHelpMenu, $idAboutItem, $idOkButton, $idCancelButton
	Local $iMsg, $sFile

	#forceref $idSeparator1
	GUICreate("PowerPoint Converter", 300, 120)

	$idFileMenu = GUICtrlCreateMenu("File")
	$idFileItem = GUICtrlCreateMenuItem("Open...", $idFileMenu)
	$idSeparator1 = GUICtrlCreateMenuItem("", $idFileMenu)
	$idHelpMenu = GUICtrlCreateMenu("?")
	$idAboutItem = GUICtrlCreateMenuItem("About", $idHelpMenu)

	$idOkButton = GUICtrlCreateButton("Convert", 50, 70, 70, 20)
	$idCancelButton = GUICtrlCreateButton("Quit", 180, 70, 70, 20)
	GUICtrlCreateLabel("Directory Selected", 10, 5, 100)
	$idTITLE = GUICtrlCreateInput("", 10, 20, 280, 20)

	GUISetState()

	While 1
		$iMsg = GUIGetMsg()

		Select
			Case $iMsg = $GUI_EVENT_CLOSE Or $iMsg = $idCancelButton
				ExitLoop

			Case $iMsg = $idFileItem
				_Pick_Directory()
				GUICtrlCreateLabel("Directory Selected", 10, 5, 100)
				$idTITLE = GUICtrlCreateInput($DIR_MESSAGE, 10, 20, 280, 20)

			Case $iMsg = $idOkButton
				if ($File_Has_Been_Selected = 1) then
				   _Convert()
				   MsgBox($MB_SYSTEMMODAL, "Done", "Conversions Complete" & @CRLF)
				   ExitLoop
				Else
				   MsgBox($MB_SYSTEMMODAL, "Error", "You need to select a directory first" & @CRLF)
				EndIf

			Case $iMsg = $idAboutItem
				MsgBox($MB_SYSTEMMODAL, "About", "         PowerPoint Transformer" & @CRLF & "Converts 4x3 Presentations into 16x9" & @CRLF & @CRLF & "  Written by Rich Alexander (2022)")
		EndSelect

	WEnd

	GUIDelete()
EndFunc

;This Function allows the user to pick a parent directory then be traverses to convert PowerPoints
Func _Pick_Directory()
	$directory = FileSelectFolder("Choose Folder to Convert Powerpoint Files", "")
		If @error then
			MsgBox($MB_SYSTEMMODAL, "Warning", "No directory was selected")
			Return;
		Else
			$File_Has_Been_Selected = 1
			$DIR_MESSAGE=$directory
		Endif

	;Create an array of files based on the parent directory you choose
	list($directory, 0)
	$listing = StringTrimRight($listing, 1)
	$listing = _StringExplode($listing, "|", 0)
	_ArraySort($listing)

EndFunc

;This function does the conversion of the PowerPoint files
Func _Convert()
	;Open Powerpoint Application
	Global $oPPT = _PPT_Open()

	;Main Loop that goes through each file and converts it from 4x3 to 16x9
	For $iCount = 0 To $Total_Files - 1
		File_Converter($listing[$iCount])
	Next

	;Close the Powerpoint Application
	_PPT_Close($oPPT)

EndFunc

;This Function recursively builds a List of all the files to be converted including sub directories
Func list($path = "", $counter = 0)
	$counter = 0
	$path &= '\'
	Local $Check_File_Type_ppt
	Local $Check_File_Type_pptx
	Local $Check_First_Char_Of_File
	Local $list_files = '', $file, $demand_file = FileFindFirstFile($path & '*')
	If $demand_file = -1 Then Return ''
		While 1
			$file = FileFindNextFile($demand_file)
				If @error Then ExitLoop
			If @extended Then
				If $counter >= 10 Then ContinueLoop
					list($path & $file, $counter + 1)
				Else
					$Check_File_Type_pptx = StringRight($file,5)
					$Check_File_Type_ppt = StringRight($file,4)
					$Check_First_Char_Of_File = StringLeft($file,1)
					; Only include PPT and PPTX.  Any hidden files with PPT or PPTX ignore as well
					if ( ($Check_File_Type_pptx = ".pptx" or $Check_File_Type_ppt = ".ppt") and $Check_First_Char_Of_File <> "." and $Check_First_Char_Of_File <> "~") Then
					$Total_Files = $Total_Files + 1
					$listing &= $path & $file & "|"
				Endif
			EndIf
		WEnd
	FileClose($demand_file)
EndFunc

;This Function converts the files from 4x3 to 16x9 and resizes the image in the file
Func File_Converter($File_Name)
	Local Const $PpSlideSizeOnScreen16x9 = 15 ;Set a variable to represent the 16x9 format
	Local $New_File_Name = $File_Name
	Local $Check_File_Type_ppt

	;Set a variable with the file path and name to be opened
	Local $sPresentation = $File_Name

	;Open up the Powerpoint file
	Local $oPresentation = _PPT_PresentationOpen($oPPT, $sPresentation, True)
		if @error then MsgBox($MB_SYSTEMMODAL, "Error", "Failed to Open PowerPoint File")

	;Change the powerpoint to a 16 x 9 format
	$oPresentation.PageSetup.SlideSize = $PpSlideSizeOnScreen16x9

	;Loop thtough all images and adjust them to the full screen size
	Local $curSlide, $curShape, $localShapeCounter
	For $curSlide In $oPPT.ActivePresentation.Slides
		$localShapeCounter=0
		For $curShape In $curSlide.Shapes
			$localShapeCounter +=1
			With $curShape
				; For the powerpoints, Object 2 is the actual sheet music
				if ($localShapeCounter=2) then
					.LockAspectRatio = False
					.Left = 0
					.Top = 0
					;.ScaleHeight(3.38, False)
					.ScaleWidth(1.33, False)
					.Rotation = 0
				EndIf
				; Objects 3 and 4 are the song number and the song title
				if ($localShapeCounter=3 or $localShapeCounter=4) then
					.LockAspectRatio = False
					if( .left < 100 ) then
						; Put song title top far left
						.left=0
					else
						; Put song number to far right
						.left=670
					EndIf
					.Top = 0
					.Rotation = 0
				EndIf
			EndWith
		Next ; curShape
	Next ; curSlide

	;Parse new File Name
	$Check_File_Type_ppt = StringRight($New_File_Name, 3)
	if ($Check_File_Type_ppt = "ppt") Then
		$New_File_Name = StringTrimRight($New_File_Name, 4)
	else
		$New_File_Name = StringTrimRight($New_File_Name, 5)
	EndIf

	;SaveAs the Powerpoint with a new filename
	$New_File_Name = $New_File_Name & "(16x9)"
	_PPT_PresentationSaveAs($oPresentation, $New_File_Name,$ppSaveAsPresentation, True)
		if @error then MsgBox($MB_SYSTEMMODAL, "Error", "Failed to Save PowerPoint File")

	;Close the new Powerpoint file
	_PPT_PresentationClose($oPresentation)
		if @error then MsgBox($MB_SYSTEMMODAL, "Error", "Failed to Close PowerPoint File")
EndFunc