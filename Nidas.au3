#Region Copyright
#cs
	Copyright (c) 2021, Sandra
#ce
#EndRegion Copyright

#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\IconPacks\iconfinder_share-square-o_1608415.ico
;#AutoIt3Wrapper_UseX64=Y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <String.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Clipboard.au3>
#include <GuiEdit.au3>
#include <ComboConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#include <WinAPIFiles.au3>
#include <InetConstants.au3>

Local $cpu_arch_type = @CPUArch
If StringRegExp($cpu_arch_type,'X64') = 1  Then
	$cek_file = FileExists(@WindowsDir & "\System32\sqlite3_x64.dll")
	If $cek_file = 0 Then
		#RequireAdmin
		FileCopy(@ScriptDir & "\sqlite3_x64.dll",@WindowsDir & "\System32\",$FC_OVERWRITE)
	EndIf
Else
	$cek_file = FileExists(@WindowsDir & "\System32\sqlite3.dll")
	If $cek_file = 0 Then
		#RequireAdmin
		FileCopy(@ScriptDir & "\sqlite3.dll",@WindowsDir & "\System32\",$FC_OVERWRITE)
	EndIf
EndIf

#Region ### START Koda GUI section ### Form=C:\Users\Sandr\Dropbox\Autoit\Form Project\hijrahdatadummy v15 simplify.kxf
$Form1_1 = GUICreate("Nidas", 346, 687, -1, -1)
GUISetBkColor(0xC0DCC0)
$Edit1 = GUICtrlCreateEdit("", 8, 128, 273, 49)
GUICtrlSetData(-1, "Edit1")
$Button1 = GUICtrlCreateButton("Save", 96, 424, 75, 57)
GUICtrlSetBkColor(-1, 0x008000)
$Input1 = GUICtrlCreateInput("", 200, 344, 137, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input2 = GUICtrlCreateInput("", 267, 312, 70, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input3 = GUICtrlCreateInput("", 184, 312, 70, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Button2 = GUICtrlCreateButton("CDatek", 286, 200, 51, 49)
$Button3 = GUICtrlCreateButton("Open File", 8, 448, 78, 33)
$Radio1 = GUICtrlCreateRadio("AO", 96, 400, 33, 17)
$Radio2 = GUICtrlCreateRadio("MO", 136, 400, 41, 17)
$Input4 = GUICtrlCreateInput("", 48, 320, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Edit2 = GUICtrlCreateEdit("", 8, 200, 273, 49)
GUICtrlSetData(-1, "Edit2")
$Input5 = GUICtrlCreateInput("", 8, 272, 161, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input6 = GUICtrlCreateInput("", 50, 296, 38, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input7 = GUICtrlCreateInput("", 131, 296, 38, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input8 = GUICtrlCreateInput("", 184, 272, 153, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Button4 = GUICtrlCreateButton("Query", 8, 400, 78, 41)
$Edit3 = GUICtrlCreateEdit("", 8, 8, 329, 97)
GUICtrlSetData(-1, "Edit3")
$Button7 = GUICtrlCreateButton("SC", 267, 368, 70, 33)
$Button8 = GUICtrlCreateButton("TOM", 184, 368, 78, 73)
$Button9 = GUICtrlCreateButton("Internet", 184, 448, 78, 33)
$Button10 = GUICtrlCreateButton("Voice", 267, 448, 70, 33)
$Button11 = GUICtrlCreateButton("ClearMO", 184, 488, 78, 33)
$Input9 = GUICtrlCreateInput("", 8, 528, 329, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Button12 = GUICtrlCreateButton("CDetails", 286, 128, 51, 49)
$Button13 = GUICtrlCreateButton("TOM asap", 267, 408, 70, 33)
$Input28 = GUICtrlCreateInput("", 48, 344, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Input29 = GUICtrlCreateInput("", 48, 368, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Label1 = GUICtrlCreateLabel("@Sandrr4", 272, 656, 65, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Incident Details", 8, 112, 77, 17)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("Datek", 8, 184, 33, 17)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("Input IP / HOSTNAME", 40, 256, 113, 17)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label5 = GUICtrlCreateLabel("IP / HOSTNAME", 216, 256, 86, 17)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label6 = GUICtrlCreateLabel("SLOT", 7, 298, 32, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label7 = GUICtrlCreateLabel("PORT", 87, 298, 34, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label8 = GUICtrlCreateLabel("CPE", 11, 320, 25, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label9 = GUICtrlCreateLabel("HSI+", 12, 344, 28, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label10 = GUICtrlCreateLabel("VOIP+", 9, 368, 35, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label11 = GUICtrlCreateLabel("Vlan HSI", 199, 296, 46, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label12 = GUICtrlCreateLabel("Vlan VOIP", 274, 296, 53, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Label13 = GUICtrlCreateLabel("ID", 183, 344, 15, 17, BitOR($SS_CENTER,$SS_CENTERIMAGE))
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
$Combo1 = GUICtrlCreateCombo(" ", 7, 496, 78, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData($Combo1, "ogp PI|ValidateWFM|Eskalasi|Double|PI", "ogp PI")
$Button5 = GUICtrlCreateButton("Copy Status", 96, 488, 75, 33)
$Edit4 = GUICtrlCreateEdit("", 8, 560, 97, 89, BitOR($GUI_SS_DEFAULT_EDIT,$ES_READONLY))
GUICtrlSetData(-1, "")
$Edit5 = GUICtrlCreateEdit("", 112, 560, 89, 89, BitOR($GUI_SS_DEFAULT_EDIT,$ES_READONLY))
GUICtrlSetData(-1, "")
$Edit6 = GUICtrlCreateEdit("", 208, 560, 129, 89, BitOR($GUI_SS_DEFAULT_EDIT,$ES_READONLY))
GUICtrlSetData(-1, "")
$Button6 = GUICtrlCreateButton("ClearInput", 267, 488, 70, 33)
$Input10 = GUICtrlCreateInput("", 208, 656, 57, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$ES_READONLY))
$Button14 = GUICtrlCreateButton("Copy SC", 8, 654, 99, 25)
$Button15 = GUICtrlCreateButton("Copy TOM", 112, 654, 91, 25)
Dim $Form1_1_AccelTable[4][2] = [["{F4}", $Button1],["{F2}", $Button2],["{F3}", $Button4],["{F1}", $Button12]]
GUISetAccelerators($Form1_1_AccelTable)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

$aBox = MsgBox(4, "Message", "Update database ?")
If $aBox = 6 Then
    $ping = Ping("www.google.com")
	If Not @error Then
	$alamat_download = "alamat_download_database"
	$nama_file = "\coba.pdf"
	$cek_ukuran_file = InetGetSize($alamat_download)
		If $cek_ukuran_file > 0 Then
			ControlSetText("Nidas", "", 25, "proses download gambar sebesar " & FormatFileSize($cek_ukuran_file))
			$hDownload = InetGet($alamat_download, @ScriptDir & $nama_file, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
			Do
				Sleep(250)
				ControlSetText("Nidas", "", 25, "downloading " & FormatFileSize($cek_ukuran_file) & " database .....")
			Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)
			Local $iBytesSize = InetGetInfo($hDownload, $INET_DOWNLOADREAD)
			ControlSetText("Nidas", "", 25, "gambar sebesar " & FormatFileSize($iBytesSize) & " sudah terdownload")
			InetClose($hDownload)
		Else
			ControlSetText("Nidas", "", 25, "Tidak dapat terhubung dengan server")
		EndIf
	Else
		ControlSetText("Nidas", "", 25, "Internet NOK, tidak dapat update database")
	EndIf
;ElseIf $aBox = 7 Then
;    MsgBox(0, "No", "No was clicked")
EndIf

DirCreate (@ScriptDir & "\Hasil\")

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_SQLite_Close()
			_SQLite_Shutdown()
			Exit

		Case $Button6 ;clear datek5t
			ControlSetText("Nidas", "", 14, "") ;ip
			ControlSetText("Nidas", "", 15, "") ;slot
			ControlSetText("Nidas", "", 16, "") ;port
			ControlSetText("Nidas", "", 12, "") ;cpe
			ControlSetText("Nidas", "", 28, "") ;inputhsi
			ControlSetText("Nidas", "", 29, "") ;inputvoip
		Case $Button4 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; Query
			Local $hQuery, $hQuery1, $hQuery2, $hQuery3
			_SQLite_Startup()
			_SQLite_Open(@ScriptDir & '\data2.db')
			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; query mecah spo
			$sText = GUICtrlRead($Edit2)
			If StringInStr($sText, "172.", 1) > 1 Then
				If StringInStr($sText, "n 1/1", 1) > 0 Then
					Call("PotSPO")
				ElseIf StringInStr($sText, "h 1/1", 1) > 0 Then
					Call("PotSPO")
				ElseIf StringInStr($sText, "1/") or StringInStr($sText, "0/") > 0 Then
					Call("PotSP")
				ElseIf StringInStr($sText, 'DEFAULT') > 1 Then
					Call("PotSPALU")
				EndIf

				If StringInStr($sText, "172.") > 0 Then
					$bArray = StringRegExp($sText, '172.(.*)', $STR_REGEXPARRAYFULLMATCH)
					;ConsoleWrite($bArray[0] & @CRLF)
					ControlSetText("Nidas", "", 14, $bArray[0])
				EndIf
			EndIf

			;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; query ip slot port
			$IP = "'" & GUICtrlRead($Input5) & "'"
			If StringInStr($IP, '1', 1) Or StringInStr($IP, 'O', 1) > 1 Then
				$Slot = GUICtrlRead($Input6)
				$Port = GUICtrlRead($Input7)


				If StringInStr($IP, "GPON") > 1 Then
					;ConsoleWrite("masuk sini")
					_SQLite_QuerySingleRow(-1, 'SELECT IP FROM dataIP WHERE HOSTNAME=' & $IP & ';', $hQuery3)
					$hasilip = "'" & $hQuery3[0] & "'"
					If StringRegExp($Slot,'^[0-9]+$',$STR_REGEXPMATCH) and StringRegExp($Port,'^[0-9]+$',$STR_REGEXPMATCH) = 1 Then
						_SQLite_QuerySingleRow(-1, 'SELECT ID FROM idsp WHERE IP=' & $hasilip & ' AND SLOT=' & $Slot & ' AND PORT=' & $Port & ';', $hQuery)
						_SQLite_QuerySingleRow(-1, 'SELECT HSI FROM dataIP WHERE IP=' & $hasilip & ';', $hQuery1)
						_SQLite_QuerySingleRow(-1, 'SELECT VOIP FROM dataIP WHERE IP=' & $hasilip & ';', $hQuery2)
						ControlSetText("Nidas", "", 5, $hQuery[0])
						ControlSetText("Nidas", "", 7, $hQuery1[0])
						ControlSetText("Nidas", "", 6, $hQuery2[0])
						ControlSetText("Nidas", "", 17, $hQuery3[0])
						$Baca_Hasil_ID = GUICtrlRead($Input1)
						If StringRegExp($Baca_Hasil_ID,'\w',$STR_REGEXPMATCH) = 0 Then
							If $Slot < 20 Then
								$hasilnya = $hQuery3[0] & ":" & $Slot
								$Cek_Missing = FileRead(@ScriptDir & "\Missing Slot.txt")
								If StringRegExp($Cek_Missing, $hasilnya) = 0 Then
									FileWrite(@ScriptDir & "\Missing Slot.txt" , $hasilnya & @CRLF)
								EndIf
							EndIf
						EndIf
					Else
						_SQLite_QuerySingleRow(-1, 'SELECT HSI FROM dataIP WHERE IP=' & $hasilip & ';', $hQuery1)
						_SQLite_QuerySingleRow(-1, 'SELECT VOIP FROM dataIP WHERE IP=' & $hasilip & ';', $hQuery2)
						ControlSetText("Nidas", "", 5, "")
						ControlSetText("Nidas", "", 7, $hQuery1[0])
						ControlSetText("Nidas", "", 6, $hQuery2[0])
						ControlSetText("Nidas", "", 17, $hQuery3[0])
					EndIf
					Else
					;ConsoleWrite("masuk situ" & @CRLF)
					_SQLite_QuerySingleRow(-1, 'SELECT HOSTNAME FROM dataIP WHERE IP=' & $IP & ';', $hQuery3)
					If StringRegExp($Slot,'^[0-9]+$',$STR_REGEXPMATCH) and StringRegExp($Port,'^[0-9]+$',$STR_REGEXPMATCH) = 1 Then
						;ConsoleWrite("masuk situ situ" & @CRLF)
						_SQLite_QuerySingleRow(-1, 'SELECT ID FROM idsp WHERE IP=' & $IP & ' AND SLOT=' & $Slot & ' AND PORT=' & $Port & ';', $hQuery)
						_SQLite_QuerySingleRow(-1, 'SELECT HSI FROM dataIP WHERE IP=' & $IP & ';', $hQuery1)
						_SQLite_QuerySingleRow(-1, 'SELECT VOIP FROM dataIP WHERE IP=' & $IP & ';', $hQuery2)
						ControlSetText("Nidas", "", 5, $hQuery[0])
						ControlSetText("Nidas", "", 7, $hQuery1[0])
						ControlSetText("Nidas", "", 6, $hQuery2[0])
						ControlSetText("Nidas", "", 17, $hQuery3[0])
						$Baca_Hasil_ID = GUICtrlRead($Input1)
						If StringRegExp($Baca_Hasil_ID,'\w',$STR_REGEXPMATCH) = 0 Then
							If $Slot < 20 Then
								$hasilnya = StringReplace($IP,"'","") & ":" & $Slot
								$Cek_Missing = FileRead(@ScriptDir & "\Missing Slot.txt")
								If StringRegExp($Cek_Missing, $hasilnya) = 0 Then
									FileWrite(@ScriptDir & "\Missing Slot.txt" , $hasilnya & @CRLF)
								EndIf
							EndIf
						EndIf
					Else
						;ConsoleWrite("masuk situ situ situ" & @CRLF)
						_SQLite_QuerySingleRow(-1, 'SELECT HSI FROM dataIP WHERE IP=' & $IP & ';', $hQuery1)
						_SQLite_QuerySingleRow(-1, 'SELECT VOIP FROM dataIP WHERE IP=' & $IP & ';', $hQuery2)
						ControlSetText("Nidas", "", 5, "")
						ControlSetText("Nidas", "", 7, $hQuery1[0])
						ControlSetText("Nidas", "", 6, $hQuery2[0])
						ControlSetText("Nidas", "", 17, $hQuery3[0])
					EndIf
				EndIf


			EndIf

			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				If StringInStr($Dataekst,@CRLF) > 1 Then
					$Dataekst = StringReplace($Dataekst,@CRLF,'')
				EndIf

				ControlSetText("Nidas", "", 19, "")
				If StringInStr($Dataekst, "Service ID is") > 0 Then
					$bebe = StringInStr($Dataekst, "Service ID is ") + 13
					$baruu = StringTrimLeft($Dataekst, $bebe)
					$baru = StringSplit($baruu, ",")
					For $i = 1 To UBound($baru) - 1
						_GUICtrlEdit_InsertText(19, $baru[$i] & @CRLF, 0)
					Next
				EndIf

				If StringInStr($Dataekst, "OSM ID") > 0 Then
					$OSM = StringRegExp($Dataekst, 'OSM ID: (.........)', $STR_REGEXPARRAYFULLMATCH)
					$OSM1 = StringMid($OSM[0], 9, 17)
					_GUICtrlEdit_InsertText(19, $OSM1 & @CRLF, 0)
				EndIf

				If StringInStr($Dataekst, "SC") > 0 Then
					$SC = StringRegExp($Dataekst, 'SC(.........)', $STR_REGEXPARRAYFULLMATCH)
					_GUICtrlEdit_InsertText(19, $SC[0] & "-" & UBound($baru) - 1 & ' layanan.' & @CRLF, 0)
				EndIf

				If StringInStr($Dataekst, "WB") > 0 Then
					$WB = StringRegExp($Dataekst, 'WB(............)', $STR_REGEXPARRAYFULLMATCH)
					_GUICtrlEdit_InsertText(19, $WB[0] & @CRLF, 0)
				EndIf

			EndIf

		Case $Button1 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; save
			Local $counter

			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst,@CRLF) > 1 Then
				$Dataekst = StringReplace($Dataekst,@CRLF,'')
			EndIf

			If StringInStr($Dataekst, "SC") > 1 Then
				If StringInStr($Dataekst, "INTERNET") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							If StringInStr($baru[$i], "INTERNET") > 0 Then
								$inet = _StringBetween($baru[$i], "_", "_INTERNET")
								$ncli = _StringBetween($baru[$i], "", "_")
								$nclidaninet = $ncli[0] & "_" & $inet[0]
							EndIf
						Next
					EndIf
				ElseIf StringInStr($Dataekst, "IPTV") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							If StringInStr($baru[$i], "IPTV") > 0 Then
								$inet = _StringBetween($baru[$i], "_", "_IPTV")
								$ncli = _StringBetween($baru[$i], "", "_")
								$nclidaninet = $ncli[0] & "_" & $inet[0]
								If StringLen($inet[0]) > 12 Then
									$yes = StringLeft($inet[0], 12)
									$nclidaninet = $ncli[0] & "_" & $yes
								Else
								EndIf
							EndIf
						Next
					EndIf
				ElseIf StringInStr($Dataekst, "VOI") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							$ncli = _StringBetween($baru[$i], "", "_")
						Next
					EndIf
				EndIf

				$IP = "'" & GUICtrlRead($Input5) & "'"
				If StringInStr($Dataekst, "SC") And StringInStr($IP, "172.") Or StringInStr($IP, "O") > 1 Then
					$ID = GUICtrlRead($Input1)
					$VHSI = GUICtrlRead($Input3)
					$VVVOIP = GUICtrlRead($Input2)
					$aDays = StringSplit($VVVOIP, "-")
					If IsArray($aDays) = 1 Then
						$VVOIP = Random($aDays[1], $aDays[2], 1)
					EndIf
					;ConsoleWrite($ID & @CRLF)
					$CPE = GUICtrlRead($Input4)

					If GUICtrlRead($Radio1) = $GUI_CHECKED Then
						$namaoutputlog = 'AO.csv'

						$hFO1 = FileOpen(@ScriptDir & "\Hasil\" & $namaoutputlog, 2)
						FileWrite($hFO1, "")
						FileClose($hFO1)
						FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
						If StringInStr($Dataekst, "Service ID is") > 0 Then
							$bebe = StringInStr($Dataekst, "Service ID is ") + 13
							$baruu = StringTrimLeft($Dataekst, $bebe)
							$baru = StringSplit($baruu, ",")
							For $i = 1 To UBound($baru) - 1
								If StringInStr($baru[$i], "INTER") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $VHSI & ',' & $baru[$i] & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								ElseIf StringInStr($baru[$i], "VOI") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $VVOIP & ',' & $baru[$i] & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								Else
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									;FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, '111,' & $baru[$i] & ',,S-Vlan,,,' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								EndIf
							Next

							$kolomsc = GUICtrlRead($Edit4)
							If StringInStr($kolomsc,'SC') = 0 Then
							_GUICtrlEdit_InsertText(45, $SC[0], 0)
							Else
							_GUICtrlEdit_InsertText(45, $SC[0] & @CRLF, 0)
							EndIf

							$kolomosmf = GUICtrlRead($Edit5)
							If StringInStr($kolomosmf,'3') = 0 Then
							_GUICtrlEdit_InsertText(46, $OSM1, 0)
							Else
							_GUICtrlEdit_InsertText(46, $OSM1 & @CRLF, 0)
							EndIf

							$kolomsc = GUICtrlRead($Edit6)
							If StringInStr($kolomsc,'SC') = 0 Then
							_GUICtrlEdit_InsertText(47, $Dataekst, 0)
							Else
							_GUICtrlEdit_InsertText(47, $Dataekst & @CRLF, 0)
							EndIf


							$counter = $counter + 1
							ControlSetText("Nidas", "", 49, $counter)

							FileWrite(@ScriptDir & "\Hasil\ Rekap " & @MDAY & @MON & @YEAR & ".txt" , "AO-" & $Dataekst & @CRLF)

							;_GUICtrlEdit_InsertText(45, $SC[0] & @CRLF, 0)
							;_GUICtrlEdit_InsertText(46, $OSM1 & @CRLF, 0)
							;_GUICtrlEdit_InsertText(47, StringReplace($addvoip,",","") & @CRLF, -1)
							;ControlSetText("Nidas", "", 45, $SC[0])
							;ControlSetText("Nidas", "", 46, $OSM1)
							;ControlSetText("Nidas", "", 47, "Data AO sudah tersimpan")
						EndIf
						ControlSetText("Nidas", "", 25, "Data AO sudah tersimpan")

					ElseIf GUICtrlRead($Radio2) = $GUI_CHECKED Then
						$namaoutputlog = 'MO.csv'

						$hFO1 = FileOpen(@ScriptDir & "\Hasil\" & $namaoutputlog)
						If StringRegExp(FileRead($hFO1),'RESOURCE_ID') = 0 Then
							$hFO1 = FileOpen(@ScriptDir & "\Hasil\" & $namaoutputlog, 2)
							FileWrite($hFO1, "")
							FileClose($hFO1)
							FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
						EndIf

						If StringInStr($Dataekst, "Service ID is") > 0 Then
							If StringInStr($Dataekst, "IPTV") And StringInStr($Dataekst, "INTERNET") > 1 Then
								If StringLen($inet[0]) > 12 Then
									$inputvoip = GUICtrlRead($Input29)
									If StringInStr($inputvoip, "03") > 0 Then
										$sVOIP = $inputvoip
									Else
										$sVOIP = InputBox("Nomor Voip", "Harap masukkan nomor voip", "")
									EndIf
									If StringInStr($sVOIP, "03") > 0 Then
										$addiptv = "," & $nclidaninet & "_IPTV+"
										$addvoip = "," & $ncli[0] & "_" & $sVOIP & "_VOICE+"
										$Dataekst1 = $Dataekst & $addvoip & $addiptv
										_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
										_GUICtrlEdit_InsertText(19, StringReplace($addvoip,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 111, "")
										ControlSetText("Nidas", "", 112, "")
									Else
										$addiptv = "," & $nclidaninet & "_IPTV+"
										$Dataekst1 = $Dataekst & $addiptv
										_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									EndIf

								Else
									$inputvoip = GUICtrlRead($Input29)
									If StringInStr($inputvoip, "03") > 0 Then
										$sVOIP = $inputvoip
									Else
										$sVOIP = InputBox("Nomor Voip", "Harap masukkan nomor voip", "")
									EndIf
									If StringInStr($sVOIP, "03") > 0 Then
										$addvoip = "," & $ncli[0] & "_" & $sVOIP & "_VOICE+"
										$Dataekst1 = $Dataekst & $addvoip
										_GUICtrlEdit_InsertText(19, StringReplace($addvoip,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									Else
										$Dataekst1 = $Dataekst
									EndIf
								EndIf
							ElseIf StringInStr($Dataekst, "VOI") And StringInStr($Dataekst, "INTERNET") > 1 Then
								$addiptv = "," & $nclidaninet & "_IPTV+"
								$Dataekst1 = $Dataekst & $addiptv
								_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
								FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
								ControlSetText("Nidas", "", 112, "")
							ElseIf StringInStr($Dataekst, "VOI") And StringInStr($Dataekst, "IPTV") > 1 Then
								If StringLen($inet[0]) > 12 Then
									$addiptv = "," & $nclidaninet & "_IPTV+"
									$addhsi = "," & $nclidaninet & "_INTERNET+"
									$Dataekst1 = $Dataekst & $addiptv & $addhsi
									_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
									_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
									ControlSetText("Nidas", "", 112, "")
								Else
									$addhsi = "," & $nclidaninet & "_INTERNET+"
									$Dataekst1 = $Dataekst & $addhsi
									_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
									ControlSetText("Nidas", "", 112, "")
								EndIf
							ElseIf StringInStr($Dataekst, "INTERNET") > 1 Then
								$inputvoip = GUICtrlRead($Input29)
								If StringInStr($inputvoip, "03") > 0 Then
									$sVOIP = $inputvoip
								Else
									$sVOIP = InputBox("Nomor Voip", "Harap masukkan nomor voip", "")
								EndIf
								If StringInStr($sVOIP, "03") > 0 Then
									$addiptv = "," & $nclidaninet & "_IPTV+"
									$addvoip = "," & $ncli[0] & "_" & $sVOIP & "_VOICE+"
									$Dataekst1 = $Dataekst & $addvoip & $addiptv
									_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
									_GUICtrlEdit_InsertText(19, StringReplace($addvoip,",","") & @CRLF, -1)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
									ControlSetText("Nidas", "", 112, "")
								Else
									$addiptv = "," & $nclidaninet & "_IPTV+"
									$Dataekst1 = $Dataekst & $addiptv
									_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
									ControlSetText("Nidas", "", 112, "")
								EndIf

							ElseIf StringInStr($Dataekst, "IPTV") > 1 Then
								If StringLen($inet[0]) > 12 Then
									$inputvoip = GUICtrlRead($Input29)
									If StringInStr($inputvoip, "03") > 0 Then
										$sVOIP = $inputvoip
									Else
										$sVOIP = InputBox("Nomor Voip", "Harap masukkan nomor voip", "")
									EndIf
									If StringInStr($sVOIP, "03") > 0 Then
										$addiptv = "," & $nclidaninet & "_IPTV+"
										$addhsi = "," & $nclidaninet & "_INTERNET+"
										$addvoip = "," & $ncli[0] & "_" & $sVOIP & "_VOICE+"
										$Dataekst1 = $Dataekst & $addhsi & $addvoip & $addiptv
										_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
										_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
										_GUICtrlEdit_InsertText(19, StringReplace($addvoip,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									Else
										$addiptv = "," & $nclidaninet & "_IPTV+"
										$addhsi = "," & $nclidaninet & "_INTERNET+"
										$Dataekst1 = $Dataekst & $addhsi & $addiptv
										_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
										_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									EndIf

								Else
									$inputvoip = GUICtrlRead($Input29)
									If StringInStr($inputvoip, "03") > 0 Then
										$sVOIP = $inputvoip
									Else
										$sVOIP = InputBox("Nomor Voip", "Harap masukkan nomor voip", "")
									EndIf
									If StringInStr($sVOIP, "03") > 0 Then
										$addvoip = "," & $ncli[0] & "_" & $sVOIP & "_VOICE+"
										$addhsi = "," & $nclidaninet & "_INTERNET+"
										$Dataekst1 = $Dataekst & $addhsi & $addvoip
										_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
										_GUICtrlEdit_InsertText(19, StringReplace($addvoip,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									Else
										$addhsi = "," & $nclidaninet & "_INTERNET+"
										$Dataekst1 = $Dataekst & $addhsi
										_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
										ControlSetText("Nidas", "", 112, "")
									EndIf
								EndIf

							Else ;VOICETOK
								$inputinet = GUICtrlRead($Input28)
								If StringInStr($inputinet, "152") > 0 Then
									$sINTET = $inputinet
								Else
									$sINTET = InputBox("Nomor Inter", "Harap masukkan nomor Internet", "")
								EndIf
								If StringInStr($sINTET, "152") > 0 Then
									$addhsi = "," & $ncli[0] & "_" & $sINTET & "_INTERNET+"
									$addiptv = "," & $ncli[0] & "_" & $sINTET & "_IPTV+"
									$Dataekst1 = $Dataekst & $addiptv & $addhsi
									_GUICtrlEdit_InsertText(19, StringReplace($addhsi,",","") & @CRLF, -1)
									_GUICtrlEdit_InsertText(19, StringReplace($addiptv,",","") & @CRLF, -1)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
									ControlSetText("Nidas", "", 112, "")
								Else
									$Dataekst1 = $Dataekst
								EndIf
							EndIf

							$bebe = StringInStr($Dataekst1, "Service ID is ") + 13
							$baruu = StringTrimLeft($Dataekst1, $bebe)
							$baru = StringSplit($baruu, ",")
							For $i = 1 To UBound($baru) - 1
								If StringInStr($baru[$i], "INTERNET+") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $ID & ',' & StringReplace($baru[$i],"+","") & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $VHSI & ',' & StringReplace($baru[$i],"+","") & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $CPE & ',' & StringReplace($baru[$i],"+","") & ',,CPE' & @CRLF)
									EndIf
								ElseIf StringInStr($baru[$i], "IPTV+") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $ID & ',' & StringReplace($baru[$i],"+","") & ',,Service_Port' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $CPE & ',' & StringReplace($baru[$i],"+","") & ',,CPE' & @CRLF)
									EndIf
								ElseIf StringInStr($baru[$i], "VOICE+") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $ID & ',' & StringReplace($baru[$i],"+","") & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $VHSI & ',' & StringReplace($baru[$i],"+","") & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $SC[0] & " tambahan " & $namaoutputlog, $CPE & ',' & StringReplace($baru[$i],"+","") & ',,CPE' & @CRLF)
									EndIf

								ElseIf StringInStr($baru[$i], "INTER") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $VHSI & ',' & $baru[$i] & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								ElseIf StringInStr($baru[$i], "VOI") > 1 Then
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $VVOIP & ',' & $baru[$i] & ',,S-Vlan' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								Else
									FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $ID & ',' & $baru[$i] & ',,Service_Port' & @CRLF)
									If $CPE > 1 Then
										FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, $CPE & ',' & $baru[$i] & ',,CPE' & @CRLF)
									EndIf
								EndIf
							Next
							$Dataekst1 = ""

							$kolomsc = GUICtrlRead($Edit4)
							If StringInStr($kolomsc,'SC') = 0 Then
							_GUICtrlEdit_InsertText(45, $SC[0], 0)
							Else
							_GUICtrlEdit_InsertText(45, $SC[0] & @CRLF, 0)
							EndIf

							$kolomosmf = GUICtrlRead($Edit5)
							If StringInStr($kolomosmf,'3') = 0 Then
							_GUICtrlEdit_InsertText(46, $OSM1, 0)
							Else
							_GUICtrlEdit_InsertText(46, $OSM1 & @CRLF, 0)
							EndIf

							$kolomsc = GUICtrlRead($Edit6)
							If StringInStr($kolomsc,'SC') = 0 Then
							_GUICtrlEdit_InsertText(47, $Dataekst, 0)
							Else
							_GUICtrlEdit_InsertText(47, $Dataekst & @CRLF, 0)
							EndIf


							$counter = $counter + 1
							ControlSetText("Nidas", "", 49, $counter)

							FileWrite(@ScriptDir & "\Hasil\ Rekap " & @MDAY & @MON & @YEAR & ".txt" , "MO-" & $Dataekst & @CRLF)
						EndIf
						ControlSetText("Nidas", "", 25, "Data MO sudah tersimpan")

					EndIf
				Else
					MsgBox(0, "", "tidak ada data untuk disave")
				EndIf
			EndIf
		Case $Button2
			;ControlSetText("Nidas", "", 5, "") ;idslot
			;ControlSetText("Nidas", "", 7, "") ;vhsi
			;ControlSetText("Nidas", "", 6, "") ;vvoip
			ControlSetText("Nidas", "", 13, "") ;datek
			;ControlSetText("Nidas", "", 14, "") ;ip
			;ControlSetText("Nidas", "", 15, "") ;slot
			;ControlSetText("Nidas", "", 16, "") ;port
			;ControlSetText("Nidas", "", 17, "") ;hostip
			;ControlSetText("Nidas", "", 25, "") ;log

		Case $Button3
			Run("explorer.exe " & @ScriptDir & "\Hasil\")
		Case $Button6
			ControlSetText("Nidas", "", 20, "")

		Case $Button7 ;copy sc
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				ClipPut($SC[0])
				ControlSetText("Nidas", "", 25, $SC[0] & "-SC sudah tersalin")
			Else
				ControlSetText("Nidas", "", 25, "Harap inputkan Incident Detailsnya. Input > Query > Klik Tombol ini")
			EndIf
		Case $Button8 ;copy tom
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				ClipPut($OSM1)
				ControlSetText("Nidas", "", 25, $OSM1 & "-TOM sudah tersalin")
			Else
				ControlSetText("Nidas", "", 25, "Harap inputkan Incident Detailsnya. Input > Query > Klik Tombol ini")
			EndIf
		Case $Button9 ;cinet
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				If StringInStr($Dataekst, "INTERNET") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							If StringInStr($baru[$i], "INTERNET") > 0 Then
								$inet = _StringBetween($baru[$i], "_", "_INTERNET")
								ClipPut($inet[0])
								ControlSetText("Nidas", "", 25, $inet[0] & "-Nomor internet sudah tersalin")
							EndIf
						Next
					EndIf
				ElseIf StringInStr($Dataekst, "IPTV") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							If StringInStr($baru[$i], "IPTV") > 0 Then
								$inet = _StringBetween($baru[$i], "_", "_IPTV")
								If StringLen($inet[0]) > 12 Then
									$yes = StringLeft($inet[0], 12)
									ClipPut($yes)
									ControlSetText("Nidas", "", 25, $yes & "-Nomor internet sudah tersalin")
								Else
									ClipPut($inet[0])
									ControlSetText("Nidas", "", 25, $inet[0] & "-Nomor internet sudah tersalin")
								EndIf
							EndIf
						Next
					EndIf
				Else
				ControlSetText("Nidas", "", 25, "Nomor internet tidak ditemukan")
				EndIf
			EndIf
		Case $Button10 ;cvoice
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				If StringInStr($Dataekst, "VOICE") > 1 Then
					If StringInStr($Dataekst, "Service ID is") > 0 Then
						$bebe = StringInStr($Dataekst, "Service ID is ") + 13
						$baruu = StringTrimLeft($Dataekst, $bebe)
						$baru = StringSplit($baruu, ",")
						For $i = 1 To UBound($baru) - 1
							If StringInStr($baru[$i], "VOICE") > 0 Then
								$voice = _StringBetween($baru[$i], "_", "_VOICE")
								ClipPut($voice[0])
								ControlSetText("Nidas", "", 25, $voice[0] & "-Nomor voip sudah tersalin")
							EndIf
						Next
					EndIf
				Else
					ControlSetText("Nidas", "", 25, "Nomor voip tidak ditemukan")
				EndIf
			Else
				ControlSetText("Nidas", "", 25, "Nomor voip tidak ditemukan")
			EndIf
		Case $Button11 ;clearmo
			$namaoutputlog = 'MO.csv'
			$hFO1 = FileOpen(@ScriptDir & "\Hasil\" & $namaoutputlog, 2)
			FileWrite($hFO1, "")
			FileClose($hFO1)
			FileWrite(@ScriptDir & "\Hasil\" & $namaoutputlog, 'RESOURCE_ID,SERVICE_NAME,TARGET_ID,CONFIG_ITEM_NAME' & @CRLF)
			ControlSetText("Nidas", "", 25, "Data MO sudah terhapus")
		Case $Button12 ;cleardetails
			ControlSetText("Nidas", "", 3, "") ;input details
			ControlSetText("Nidas", "", 112, "") ;input voip
			ControlSetText("Nidas", "", 12, "") ;cpe
		Case $Button13 ;copy tomasp
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				ClipPut($OSM1 & "%")
				ControlSetText("Nidas", "", 25, "TOM sudah tersalin")
			Else
				ControlSetText("Nidas", "", 25, "Harap inputkan Incident Detailsnya. Input > Query > Klik Tombol ini")
			EndIf
		Case $Button5
			$Dataekst = GUICtrlRead($Edit1)
			If StringInStr($Dataekst, "SC") > 1 Then
				$sComboRead = GUICtrlRead($Combo1)
				ClipPut($SC[0] & " / SC-" & StringRight($SC[0], 9) & " " & $sComboRead)
				ControlSetText("Nidas", "", 25, "status SC sudah tersalin " & $sComboRead)
			Else
				ControlSetText("Nidas", "", 25, "Harap inputkan Incident Detailsnya. Input > Query > Klik Tombol ini")
			EndIf


		Case $Button14
			ClipPut(GUICtrlRead($Edit4))
			ControlSetText("Nidas", "", 25, "SC sudah tersalin")
		Case $Button15
			ClipPut(GUICtrlRead($Edit5))
			ControlSetText("Nidas", "", 25, "TOM sudah tersalin")
	EndSwitch
WEnd

Func PotSP()
	If StringInStr($sText, "hwi") > 0 Then
		$slotprot = _StringBetween($sText, "0/", @CR)
		$yesbisa = StringSplit($slotprot[0],"/")
		ControlSetText("Nidas", "", 16, $yesbisa[2]) ;PORT
		ControlSetText("Nidas", "", 15, $yesbisa[1]) ;SLOT
	Else
	$gege = _StringBetween($sText, "1/", @CR)
	$hehey = StringSplit($gege[0], "/")
	If StringInStr($gege[0], ":") = 0 Then ;==== ALU
		;_GUICtrlEdit_InsertText(25, "PORT : " & $hehey[3] & @CRLF, 0)
		ControlSetText("Nidas", "", 16, $hehey[3])
		;_GUICtrlEdit_InsertText(25, "SLOT : " & $hehey[2] & @CRLF, 0)
		ControlSetText("Nidas", "", 15, $hehey[2])
		;_GUICtrlEdit_InsertText(25, "OLT ALU" & @CRLF, 0)
	Else ;=== ZTE

		$hihiy = StringRegExpReplace($gege[0], ":", "/")
		$hahay = StringSplit($hihiy, "/")
		;_GUICtrlEdit_InsertText(25, "PORT : " & $hahay[2] & @CRLF, 0)
		ControlSetText("Nidas", "", 16, $hahay[2])
		;_GUICtrlEdit_InsertText(25, "SLOT : " & $hahay[1] & @CRLF, 0)
		ControlSetText("Nidas", "", 15, $hahay[1])
		;_GUICtrlEdit_InsertText(25, "OLT ZTE " & @CRLF, 0)
	EndIf
	EndIf
EndFunc   ;==>PotSP

Func PotSPO()
	$beb = StringInStr($sText, " 1/1") + 4
	$barb = StringTrimLeft($sText, $beb)
	$kanan = StringRight($barb, 5)
	If StringInStr($kanan, "/14/1") = 1 Then
		$barbb = StringTrimRight($barb, 5)
	Else
		$barbb = StringTrimRight($barb, 4)
	EndIf
	$barbbb = StringSplit($barbb, "/")


	If StringInStr($barbbb[3], '0') = 1 Then
		$potonk = StringReplace($barbbb[3], "0", "")
		;ConsoleWrite($potong & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Onu Id : " & $potonk & @CRLF, 0)
		Global $Onu = $potonk + 0
	Else
		;ConsoleWrite($barbbb[2] & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Onu Id : " & $barbbb[3] & @CRLF, 0)
		Global $Onu = $barbbb[3] + 0
	EndIf

	If StringInStr($barbbb[2], '0') = 1 Then
		$potong = StringReplace($barbbb[2], "0", "")
		;ConsoleWrite($potong & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Port : " & $potong & @CRLF, 0)
		Global $Port = $potong + 0
		ControlSetText("Nidas", "", 16, $potong) ;port
	Else
		;ConsoleWrite($barbbb[2] & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Port : " & $barbbb[2] & @CRLF, 0)
		Global $Port = $barbbb[2] + 0
		ControlSetText("Nidas", "", 16, $barbbb[2]) ;port
	EndIf

	If StringInStr($barbbb[1], '0') = 1 Then
		$poton = StringReplace($barbbb[1], "0", "")
		;ConsoleWrite($poton & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Slot : " & $poton & @CRLF, 0)
		Global $Slot = $poton + 0
		ControlSetText("Nidas", "", 15, $poton) ;slot
	Else
		;ConsoleWrite($barbbb[1] & @CRLF)
		;_GUICtrlEdit_InsertText(25, "Slot : " & $barbbb[1] & @CRLF, 0)
		Global $Slot = $barbbb[1] + 0
		ControlSetText("Nidas", "", 15, $barbbb[1]) ;slot
	EndIf
	$gg = _StringBetween($sText, "GPON", " ")
	Global $iniajaditelnet = $gg[0]
	;_GUICtrlEdit_InsertText(25, "GPON" & $gg[0] & @CRLF, 0)
	;ControlSetText("Provi", "", 7, "GPON" & $gg[0])
EndFunc   ;==>PotSPO

Func SPO2()
	$aText = StringSplit($sText, Chr(9), 1)
	Global $Slot = $aText[3]
	Global $Port = $aText[4]
	Global $Onu = $aText[5]
	Global $iniajaditelnet = $aText[2]
EndFunc   ;==>SPO2

Func PotSPALU()
	$hasil = _StringBetween($sText,'1/1',' ')
	$hasil = StringSplit($hasil[0],'/')

	ControlSetText("Nidas", "", 16, $hasil[3]) ;PORT
	ControlSetText("Nidas", "", 15, $hasil[2]) ;SLOT
EndFunc

Func FormatFileSize($iSizeInBytes)

  Local $iSize, $i

  $iSize = $iSizeInBytes
  $i = 0
  While $iSize >= 1024 ;split value, you can set lower number like 900 for earlier switch (0.98 MB)
    $iSize /= 1024 ;this must be always 1024
    $i += 1
  WEnd

  $iSize = Round($iSize, 1) & " "; enter how many decimals you want, here it's '1'

  Switch $i
    Case 1
      $iSize = $iSize & "k"
    Case 2
      $iSize = $iSize & "M"
    Case 3
      $iSize = $iSize & "G"
    Case 4
      $iSize = $iSize & "T"
  EndSwitch

  Return $iSize & "B"

EndFunc
