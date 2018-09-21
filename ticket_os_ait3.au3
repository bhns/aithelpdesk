#CS GOD is love
 * Copyright (c) 2018 BHNS
 * Email: bhns@bhns.com.br
 * Website: www.bhns.com.br
 *          
 *
 * This file is part of SRH-Online.
 *
 * SRH-Online is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
#CE
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ait_new.ico
#AutoIt3Wrapper_Outfile=AIT HelpDesk®.exe
#AutoIt3Wrapper_Res_Comment=AIT HelpDesk ® Visite  www.bhns.com.br
#AutoIt3Wrapper_Res_Description=AIT HelpDesk ® Utilitario de suporte a base de tickets e suporte remoto.
#AutoIt3Wrapper_Res_Fileversion=0.3.0.93
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=BHns Belo Horizonte Network Solutions
#AutoIt3Wrapper_Res_Language=1046
#AutoIt3Wrapper_Res_requestedExecutionLevel=highestAvailable
#AutoIt3Wrapper_Res_Field=Version|0.3
#AutoIt3Wrapper_Res_Field=Build | 2013.08.30
#AutoIt3Wrapper_Res_Field=Coded by | BHNS www.bhns.com.br Cleber Antonio 31 8829-6064
#AutoIt3Wrapper_Res_Field=Compile date|%longdate% %time%
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Run_AU3Check=n
#Au3Stripper_Parameters=/sf /sv /om /cs=0 /cn=0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#Obfuscator_Parameters=/sf /sv /om /cs=0 /cn=0
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <array.au3>
#include<file.au3>
#include <String.au3>
#include <Date.au3>
#include <ScreenCapture.au3>
#include <GDIPlus.au3>
#include <File.au3>
#include "lib\base64.au3"
#include "lib\AD\AD.au3"
#include "lib\WinHttp.au3"
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
AutoItSetOption("WinTitleMatchMode", 4)

$Debug_TAB = False ; Check ClassName being passed to functions, set to True and use a handle to another control to see it work
$version = "AIT HelpDesk ® v" & StringLeft(FileGetVersion(@ScriptName), 5) & "  -  Process Manager "
$Debug_LV = False; Check ClassName being passed to ListView functions, set to True and use a handle to another control to see it work

$about = "About " & $version
$MainLabel = $version;& @LF & "Portable"
$CopyRLabel = "Copyright © 2013 AIT HelpDesk ®" & @LF & "All rights reserved."
$NameURL1 = "HomePage"
$URL1 = "http://AIT HelpDesk ®"
$NameURL2 = "Email Developer"
$URL2 = "mailto:clebe@live.com"
$NameURL3 = ""
$URL3 = "http://"
$LinkColor = 0x0000FF
$BkColor = 0xAEC0FF
$hWnd = WinGetHandle(WinGetTitle(""))
$Debug_TAB = False ; Check ClassName being passed to functions, set to True and use a handle to another control to see it work
$version = "AIT HelpDesk ® V" & StringLeft(FileGetVersion(@ScriptName), 5) & "  -  Process Manager "
$Debug_LV = False; Check ClassName being passed to ListView functions, set to True and use a handle to another control to see it work


Local $random_n = Random(111111, 999999, 1)
Local $random_a = 0
Local $ticket_full_size = 0
Local $json_attachments = 0
Local $json_attachments_full
Local $folder_usernetwork = @UserProfileDir & "\AppData\Roaming\Microsoft\Network\Connections"
Local $folder_aithelpdesk = @UserProfileDir & "\.AitHelpDesk"
Local $folder_convicts = $folder_aithelpdesk & "\Convites\"
Local $folder_pictures = $folder_aithelpdesk & "\Pictures\"
Local $folder_UTIL = $folder_aithelpdesk & "\util\"
Local $last_tickets = $folder_aithelpdesk & "\Tickets\"
Local $folder_config = $folder_aithelpdesk & "\config\"
Local $tiket_img_dir = $last_tickets & $random_n & "\img\"
Local $last_tiket_reports = $last_tickets & $random_n & "\Tickets\"
$FILE_4 = $folder_pictures & "bg4.jpg" ;================
$FILE_1 = $folder_pictures & "bg1.jpg";================
$FILE_6 = $folder_pictures & "bg6.jpg";================
$FILE_10 = $folder_UTIL & "PLINK.EXE";================
$Plink_File = $FILE_10
Local $folder_pbk = $folder_usernetwork & "\Pbk\";pra criar pasta
Local $folder_bkp_pbk = $folder_usernetwork & "\old_Pbk\";pra criar pasta
HotKeySet("{F6}", "_sair")
HotKeySet("{F7}", "_About")
Local $arquivo
Local $nomeComputador, $usuarioLogado, $sistemaOperacional, $servicePack, $espacoLivre, $ipAddress, $data, $horario, $data2, $horario2
Local $inputMensagem, $inputNome, $inputEmail, $inputBreveDescr, $file_, $diretorio
Local $ticketID, $imagemSalva = 0
Local $arrayPrtScreen[1], $aPrtScreenSize = 1

; G:\xampp_old\htdocs\ticket\scp\anexos
TrayTip("AIT HelpDesk ®", "Iniciando... Bibliotecas", 5, 1)
;Sleep(1000)



Local $aUser, $g_menber_user, $g_menber_Comp, $status_remot, $random_rdp, $file_rdp, $random_KEY


; &" "& @CPUArch ;& MemGetStats ;
$data = ("" & @MDAY & "/" & @MON & "/" & @YEAR)
$horario = ("" & @HOUR & ":" & @MIN & ":" & @SEC)



$nomeComputador = @ComputerName
$usuarioLogado = @UserName
$sistemaOperacional = @OSVersion & " " & @OSType
$servicePack = @OSServicePack
$espacoLivre = DriveSpaceFree("C:\")
$ipAddress = @IPAddress1
$date_sql = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
$source_form = "Web"
$source_form1 = "AIT HelpDesk ®"
Local $ref_id_ = "R"

;Cria arquivo de configuração caso ele não exista

Local $remote_vnc_file = @ScriptDir & "\ait_helpdesk_vnc.exe";pra criar pasta

Local $sFldr = @ScriptDir & "\Config";pra criar pasta
;Local $sFldr = "\Config\";pra criar pasta

If DirGetSize($sFldr) = -1 Then
	TrayTip("AIT HelpDesk ®", "Configurando pastas de configurações", 5, 1)
	Sleep(1000)
	DirCreate($sFldr)
	IniWrite($sFldr & "\config.ini", "ENTERPRISE", "NAME", "  BHNS")
	IniWrite($sFldr & "\config.ini", "API", "HOST", "localhost")
	IniWrite($sFldr & "\config.ini", "API", "API-LINK", "osticket/api/tickets.json")
	IniWrite($sFldr & "\config.ini", "API", "KEY", "ED4FBB4F222AF67CBB5D65225E483569")
	IniWrite($sFldr & "\config.ini", "API", "FORCE-IP", "189.32.32.1")
	IniWrite($sFldr & "\config.ini", "API", "WEB-FRONT", "localhost/osticket/")
	IniWrite($sFldr & "\config.ini", "API", "IMG-AITHELPDESK", "http://192.168.0.100/upload/images/aithelpdesk/") ;
	IniWrite($sFldr & "\config.ini", "REPORT-EMAIL", "CLIENTE", "0")
	IniWrite($sFldr & "\config.ini", "REPORT-EMAIL", "API-DESK", "0")
	IniWrite($sFldr & "\config.ini", "FILE", "TRUST", "GbpSv.exe;AIT HelpDesk®.exe;igfxtray.exe;hkcmd.exe;igfxpers.exe;igfxsrvc.exe;iTunesHelper.exe;jusched.exe;cmd.exe;winbox.exe;firefox.exe;notepad++.exe;plugin-container.exe;FlashPlayerPlugin_15_0_0_152.exe;dbus-daemon.exe;kleopatra.exe;gpg-agent.exe;scdaemon.exe;TeamViewer.exe;iTunes.exe;AppleMobileDeviceHelper.exe;distnoted.exe;EXCEL.EXE;OUTLOOK.EXE;POWERPNT.EXE;WINWORD.EXE;wmplayer.exe;SoftwareUpdate.exe;ehshell.exe;AutoIt3.exe")
	IniWrite($sFldr & "\config.ini", "FILE", "WINDOWS", "MpCmdRun.exe;conhost.exe;audiodg.exe;fsquirt.exe;wordpad.exe;SoundRecorder.exe;msra.exe;SearchProtocolHost.exe;SearchFilterHost.exe;WmiPrvSE.exe;dwm.exe;SearchIndexer.exe;lsm.exe;mdm.exe;wmpnetwk.exe;System Idle Process;System;smss.exe;csrss.exe;wininit.exe;lsass.exe;svchost.exe;winlogon.exe;spoolsv.exe;services.exe;notepad.exe;taskmgr.exe;iexplore.exe;calc.exe;explorer.exe;splwow64.exe;taskhost.exe;sidebar.exe;conhost.exe;mspaint.exe;mstsc.exe;dllhost.exe")
	IniWrite($sFldr & "\config.ini", "FILE", "ALERT", "ibmpmsvc.exe;AppleMobileDeviceService.exe;mDNSResponder.exe;dirmngr.exe;TeamViewer_Service.exe;igfxext.exe;igfxsrvc.exe;iPodService.exe;xampp-control.exe;mysqld.exe;httpd.exe;tv_w32.exe;tv_x64.exe")
	IniWrite($sFldr & "\config.ini", "FILE", "DANGER", "")
	IniWrite($sFldr & "\config.ini", "WORKSTATION", "REPORT-GROUP", "1;2;3;4")
	IniWrite($sFldr & "\config.ini", "TICKET-SIZE", "ANEXO-SIZE-KB", "2048")
	IniWrite($sFldr & "\config.ini", "TICKET-SIZE", "TICKET-SIZE-KB", "10240")
	IniWrite($sFldr & "\config.ini", "DISK-WARNING", "WARNING-SIZE-MB", "15000")
	IniWrite($sFldr & "\config.ini", "DISK-WARNING", "DANGER-SIZE-MB", "1000")
	IniWrite($sFldr & "\config.ini", "USERVPN", "HOST", "HOST SERVER VPN")
	IniWrite($sFldr & "\config.ini", "USERVPN", "EMAIL", "LOGIN VPN")
	IniWrite($sFldr & "\config.ini", "USERVPN", "PASSWD", "SENHA VPN")
	IniWrite($sFldr & "\config.ini", "EMAIL", "SERVER", "smtp.gmail.com")
	IniWrite($sFldr & "\config.ini", "EMAIL", "PORTA", "465")
	IniWrite($sFldr & "\config.ini", "EMAIL", "SSL", "1")
	IniWrite($sFldr & "\config.ini", "EMAIL", "USUARIO", "email@email.com")
	IniWrite($sFldr & "\config.ini", "EMAIL", "SENHA", "XXXXXX")
	IniWrite($sFldr & "\config.ini", "PERMISSAO-ADM", "FORCE", "1")
	IniWrite($sFldr & "\config.ini", "PERMISSAO-ADM", "DOMINIO", "bhns")
	IniWrite($sFldr & "\config.ini", "PERMISSAO-ADM", "USER", "administrator")
	IniWrite($sFldr & "\config.ini", "PERMISSAO-ADM", "PASSWORD", "")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_1", "Suporte")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_2", "Sistemas")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_3", "RH")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_4", "Site")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_5", "Logistica")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_6", "Estoque")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_7", "Financeiro")
	IniWrite($sFldr & "\config.ini", "Prioridade", "NV1", "Ticket Normal")
	IniWrite($sFldr & "\config.ini", "Prioridade", "NV2", "Ticket Urgente")
	IniWrite($sFldr & "\config.ini", "Prioridade", "NV3", "Remoto Light")
	IniWrite($sFldr & "\config.ini", "Prioridade", "NV4", "Remoto RDP")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_1", "1")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_2", "2")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_3", "3")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_4", "4")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_5", "5")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_6", "6")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_ID_7", "1")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_1", "1")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_2", "2")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_3", "3")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_4", "4")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_5", "5")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_6", "6")
	IniWrite($sFldr & "\config.ini", "DP", "DPN_TP_7", "1")
	TrayTip("AIT HelpDesk ®", "Não se esqueça de configurar servidor e banco de dados" & @LF & " no arquivo config.ini " & @LF & "configurações padrões aplicadas", 5, 1)
	Sleep(3000)

EndIf




; pastas padrão do sistema

If DirGetSize($folder_aithelpdesk) = -1 Then
	DirCreate($folder_aithelpdesk)
	DirCreate($folder_convicts)
	DirCreate($last_tickets)

Else
	;DirCreate($folder_aithelpdesk)
EndIf


If DirGetSize($folder_convicts) = -1 Then
	DirCreate($folder_convicts)
Else
	;DirCreate($folder_aithelpdesk)
EndIf

If DirGetSize($last_tickets) = -1 Then
	DirCreate($last_tickets)
Else
	;DirCreate($folder_aithelpdesk)
EndIf

If DirGetSize($folder_config) = -1 Then
	DirCreate($folder_config)
	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "LAST", "0")
	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "BACKUPFILE", "0")


Else
	;DirCreate($folder_aithelpdesk)
EndIf

If DirGetSize($tiket_img_dir) = -1 Then
	DirCreate($tiket_img_dir)

Else
	;DirCreate($folder_aithelpdesk)
EndIf





;=======================PICTURES===================

If DirGetSize($folder_pictures) = -1 Then
	DirCreate($folder_pictures)
	_STRING_1(1)
	_STRING_4(1)
	_STRING_6(1)
	TrayTip("AIT HelpDesk ®", "Configurando pastas locais", 5, 1)
	Sleep(1000)
ElseIf Not FileExists($FILE_1) Then
	TrayTip("AIT HelpDesk ®", "Recuperando arquivos danificados", 5, 1)
	Sleep(1000)
	_STRING_1(1)
	_STRING_4(1)
	_STRING_6(1)
ElseIf Not FileExists($FILE_4) Then
	TrayTip("AIT HelpDesk ®", "Recuperando arquivos danificados", 5, 1)
	Sleep(1000)
	_STRING_1(1)
	_STRING_4(1)
	_STRING_6(1)
ElseIf Not FileExists($FILE_6) Then
	TrayTip("AIT HelpDesk ®", "Recuperando arquivos danificados", 5, 1)
	Sleep(1000)
	_STRING_1(1)
	_STRING_4(1)
	_STRING_6(1)
Else


EndIf



If DirGetSize($last_tiket_reports) = -1 Then
	DirCreate($last_tiket_reports)
Else
	;DirCreate($folder_aithelpdesk)
EndIf





$session_last = IniRead($folder_config & "aithelpdesk.ini", "SESSION", "BACKUPFILE", "")
$session_backup = IniRead($folder_config & "aithelpdesk.ini", "SESSION", "LAST", "")


;If DirGetSize($sFldr1) = -1 Then
;DirCreate(  $sFldr2)
;Carrega os dados do arquivo de configuração

$empresa = IniRead($sFldr & "\config.ini", "ENTERPRISE", "NAME", "")
$api_host = IniRead($sFldr & "\config.ini", "API", "HOST", "")
$api_link = IniRead($sFldr & "\config.ini", "API", "API-LINK", "")
$api_key = IniRead($sFldr & "\config.ini", "API", "KEY", "")
$api_ip = IniRead($sFldr & "\config.ini", "API", "FORCE-IP", "")
$api_web_front = IniRead($sFldr & "\config.ini", "API", "WEB-FRONT", "")
$api_web_img = IniRead($sFldr & "\config.ini", "API", "IMG-AITHELPDESK", "")
$file_trust = IniRead($sFldr & "\config.ini", "FILE", "TRUST", "") ;$file_trust $file_windows $file_alert $file_danger
$file_windows = IniRead($sFldr & "\config.ini", "FILE", "WINDOWS", "")
$file_alert = IniRead($sFldr & "\config.ini", "FILE", "ALERT", "")
$file_danger = IniRead($sFldr & "\config.ini", "FILE", "DANGER", "")
$report_workstation = IniRead($sFldr & "\config.ini", "WORKSTATION", "REPORT-GROUP", "")
$report_user_mail = IniRead($sFldr & "\config.ini", "REPORT-EMAIL", "CLIENTE", "")
$report_api_mail = IniRead($sFldr & "\config.ini", "REPORT-EMAIL", "API-DESK", "")
$get_ticket_anexo_size = IniRead($sFldr & "\config.ini", "TICKET-SIZE", "ANEXO-SIZE-KB", "")
$get_ticket_ticket_size = IniRead($sFldr & "\config.ini", "TICKET-SIZE", "TICKET-SIZE-KB", "")
$disk_warning_size = IniRead($sFldr & "\config.ini", "DISK-WARNING", "WARNING-SIZE-MB", "")
$disk_critical_size = IniRead($sFldr & "\config.ini", "DISK-WARNING", "DANGER-SIZE-MB", "")
$host_vpn = IniRead($sFldr & "\config.ini", "USERVPN", "HOST", "")
$email_vpn = IniRead($sFldr & "\config.ini", "USERVPN", "EMAIL", "")
$passwd_vpn = IniRead($sFldr & "\config.ini", "USERVPN", "PASSWD", "")
$HOST_Server = IniRead($sFldr & "\config.ini", "EMAIL", "SERVER", "")
$emailPORT = IniRead($sFldr & "\config.ini", "EMAIL", "PORTA", "")
$emailSSL = IniRead($sFldr & "\config.ini", "EMAIL", "SSL", "")
$emailServer = IniRead($sFldr & "\config.ini", "EMAIL", "USUARIO", "")
$emailSenha = IniRead($sFldr & "\config.ini", "EMAIL", "SENHA", "")
$adm_force = IniRead($sFldr & "\config.ini", "PERMISSAO-ADM", "FORCE", "")
$adm_domain = IniRead($sFldr & "\config.ini", "PERMISSAO-ADM", "DOMINIO", "")
$adm_user = IniRead($sFldr & "\config.ini", "PERMISSAO-ADM", "USER", "")
$adm_pass = IniRead($sFldr & "\config.ini", "PERMISSAO-ADM", "PASSWORD", "")
$name_DPN_1 = IniRead($sFldr & "\config.ini", "DP", "DPN_1", "")
$name_DPN_2 = IniRead($sFldr & "\config.ini", "DP", "DPN_2", "")
$name_DPN_3 = IniRead($sFldr & "\config.ini", "DP", "DPN_3", "")
$name_DPN_4 = IniRead($sFldr & "\config.ini", "DP", "DPN_4", "")
$name_DPN_5 = IniRead($sFldr & "\config.ini", "DP", "DPN_5", "")
$name_DPN_6 = IniRead($sFldr & "\config.ini", "DP", "DPN_6", "")
$name_DPN_7 = IniRead($sFldr & "\config.ini", "DP", "DPN_7", "")
$prioridade_id1 = IniRead($sFldr & "\config.ini", "Prioridade", "NV1", "")
$prioridade_id2 = IniRead($sFldr & "\config.ini", "Prioridade", "NV2", "")
$prioridade_id3 = IniRead($sFldr & "\config.ini", "Prioridade", "NV3", "")
$prioridade_id4 = IniRead($sFldr & "\config.ini", "Prioridade", "NV4", "")
$id_DPN_1 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_1", "")
$id_DPN_2 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_2", "")
$id_DPN_3 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_3", "")
$id_DPN_4 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_4", "")
$id_DPN_5 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_5", "")
$id_DPN_6 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_6", "")
$id_DPN_7 = IniRead($sFldr & "\config.ini", "DP", "DPN_ID_7", "")
$tpid_DPN_1 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_1", "")
$tpid_DPN_2 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_2", "")
$tpid_DPN_3 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_3", "")
$tpid_DPN_4 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_4", "")
$tpid_DPN_5 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_5", "")
$tpid_DPN_6 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_6", "")
$tpid_DPN_7 = IniRead($sFldr & "\config.ini", "DP", "DPN_TP_7", "")

;Send json report
If $report_api_mail = 0 Then
	$report_api_mail = "true"
Else
	$report_api_mail = "false"
EndIf



$get_ticket_anexo_size = $get_ticket_anexo_size * 1024 ; Converter byte en kbyte MB 1048576
$get_ticket_ticket_size = $get_ticket_ticket_size * 1024 ; Converter byte en kbyte

$host = "192.168.0.1"
$usr = "admin"
$pass = "senha"
$port = "22"

;EDIT THESE LINES
$rasName = "AITHELPDESK"
$rasUser = $email_vpn
$rasPass = $passwd_vpn
$destHost = $host_vpn
;Run("rasdial "&$rasName&" "&$rasUser&" "&$rasPass,"",@SW_DISABLE)
;====================PREVILEGIOS ADMINISTRATIVOS=================

If $adm_domain = "" Then

	$adm_domain = @ComputerName

EndIf

If $adm_user = "" Then

	$adm_user = @UserName

EndIf

Local $sUserName = $adm_user
Local $sPassword = $adm_pass
Local $sDomains_Work = $adm_domain


;==================OTIMIZAÇÃO PARA CARREGA DADOS =============
If FileExists($folder_aithelpdesk & "\profile.ini") Then

	TrayTip("AIT HelpDesk ®", "Importando dados aguarde...", 5, 1)
	

Else

	TrayTip("Importando Dados", "Active Directory", 5, 1)


	_AD_Open()
	;Carrega o e-mail do usuário
	$inputEmail = _AD_GetObjectAttribute(@UserName, "mail")
	;$g_menber_user = _AD_GetObjectClass(@UserName) & @CRLF & _
	;$g_menber_Comp = _AD_GetObjectClass(@ComputerName & "$"))
	$aProperties = _AD_GetObjectProperties(@UserName, "displayname,distinguishedName")
	$aProperties2 = _AD_GetObjectProperties(@ComputerName & "$")
	$aUser = _AD_GetUserGroups(@UserName)

	_AD_Close()

	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "USUARIO", $usuarioLogado)
	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EMAIL", $inputEmail)
	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "PHONE", "")
	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EXTENSION", "")
	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_TICKET", "")
	IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_ASSUNTO", "")


EndIf


$NAME_USER = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "USUARIO", "")

$NAME_EMAIL = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "EMAIL", "")

$NAME_PHONE = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "PHONE", "")

$NAME_EXTENSION = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "EXTENSION", "")

$NAME_LASTTICKET = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "LAST_TICKET", "")

$NAME_ASSUNTO = IniRead($folder_aithelpdesk & "\profile.ini", "USER", "LAST_ASSUNTO", "")
$inputEmail = $NAME_EMAIL
$ano_convert = @YEAR - 2000
;MsgBox(64, "ano convertido", $ano_convert)

;Cria o formulário

#Region ### START Koda GUI section ### Form=g:\formulario\source_form\form2.kxf
$Form1_1 = GUICreate("AIT HelpDesk ®", 412, 690, -1, -1)
GUISetBkColor(0xA6CAF0)
$Computer = GUICtrlCreateGroup("Identificação Local", 48, 576, 313, 49)
$Computador = GUICtrlCreateLabel("Nome :", 63, 597, 38, 17)
$_computer = GUICtrlCreateInput($nomeComputador, 117, 594, 89, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER, $ES_UPPERCASE, $ES_READONLY))
GUICtrlSetBkColor(-1, 0xFFFBF0)
$EndIP = GUICtrlCreateLabel("IP :", 225, 597, 20, 17)
GUICtrlCreateInput($ipAddress, 248, 593, 89, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER, $ES_UPPERCASE, $ES_READONLY))
GUICtrlSetBkColor(-1, 0xFFFBF0)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$solicitante = GUICtrlCreateGroup("Dados Solicitante", 48, 64, 313, 105)
$nome = GUICtrlCreateLabel("Nome :", 56, 88, 38, 17)
$formNome = GUICtrlCreateInput($NAME_USER, 104, 84, 233, 21)
GUICtrlSetLimit(-1, 40)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$Email = GUICtrlCreateLabel("Email :", 56, 111, 35, 17)
$formEmail = GUICtrlCreateInput($inputEmail, 104, 110, 233, 21)
GUICtrlSetLimit(-1, 40)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$telefone = GUICtrlCreateLabel("Tel. :", 56, 139, 28, 17)
$input_telefone = GUICtrlCreateInput($NAME_PHONE, 104, 136, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_UPPERCASE, $ES_NUMBER))
GUICtrlSetLimit(-1, 11)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$ramal = GUICtrlCreateLabel("Ramal :", 222, 139, 40, 17)
$Input_ramal = GUICtrlCreateInput($NAME_EXTENSION, 264, 136, 73, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_UPPERCASE, $ES_NUMBER))
GUICtrlSetLimit(-1, 11)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$suporte = GUICtrlCreateGroup("Dados da Solicitação", 48, 176, 313, 81)
$setor = GUICtrlCreateLabel("Setor :", 58, 201, 35, 17)
$formDept = GUICtrlCreateCombo($name_DPN_1, 104, 197, 97, 25, $CBS_DROPDOWN)
GUICtrlSetData(-1, $name_DPN_2 & "|" & $name_DPN_3 & "|" & $name_DPN_4 & "|" & $name_DPN_5 & "|" & $name_DPN_6 & "|" & $name_DPN_7, $name_DPN_1)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$n_tipo = GUICtrlCreateLabel("Tipo :", 202, 201, 31, 17)
$tipo = GUICtrlCreateCombo($prioridade_id1, 232, 197, 105, 25, $CBS_DROPDOWN)
GUICtrlSetData(-1, $prioridade_id2 & "|" & $prioridade_id3 & "|" & $prioridade_id4, $prioridade_id1)
GUICtrlSetBkColor(-1, 0xFFFBF0)
$solicitacao = GUICtrlCreateLabel("Assunto:", 56, 225, 45, 17)
$formBreveDescr = GUICtrlCreateInput("", 104, 221, 233, 21)
GUICtrlSetLimit(-1, 40)
GUICtrlSetBkColor(-1, 0xFFFBF0)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$botaoAnexar = GUICtrlCreateButton("Anexar arquivo", 100, 280, 87, 25)
$botaoPrtScreen = GUICtrlCreateButton("Print Screen", 188, 280, 75, 25)
$botaoVisualizar = GUICtrlCreateButton("Vizualizar", 264, 280, 75, 25)
$Detalhes = GUICtrlCreateGroup("Detalhes da Solicitação", 48, 264, 313, 305)
$formMensagem = GUICtrlCreateEdit("", 64, 312, 281, 161)
GUICtrlSetData(-1, "")
GUICtrlSetBkColor(-1, 0xFFFBF0)
$botaoEnviar = GUICtrlCreateButton("Enviar", 110, 502, 75, 40)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$botaoCancelar = GUICtrlCreateButton("Sair", 230, 502, 75, 40)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$info_detalhes = GUICtrlCreateLabel("Favor inserir detalhes da solicitação.", 120, 480, 175, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$StatusBar1 = _GUICtrlStatusBar_Create($Form1_1)
$Pic1 = GUICtrlCreatePic($FILE_1, 124, 1, 164, 60)
$Pic2 = GUICtrlCreatePic("bg2.jpg", 48, 633, 100, 28)
$Pic3 = GUICtrlCreatePic("bg3.jpg", 155, 633, 100, 28)
$Pic4 = GUICtrlCreatePic($FILE_4, 261, 633, 100, 28)
$Pic6 = GUICtrlCreatePic($FILE_6, 2, 2, 44, 28)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$bt_ferrametas = GUICtrlCreateButton("Loading Ticket", 336, 0, 75, 25)
$curtir_face = GUICtrlCreateButton("Curtir", 2, 30, 44, 25)
$Pic5 = GUICtrlCreatePic("", 0, 71, 47, 585)
$Pic7 = GUICtrlCreatePic("", 361, 30, 47, 625)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Persisted', @ScriptDir & '\AIT HelpDesk®.exe', 'REG_DWORD', 1)
FileDelete($folder_convicts & "*.*") ; Deleta todos os convites antigos na maquina

If $session_last = 1 Then

	disconect_VPN()

EndIf



TrayTip("AIT HelpDesk ®", "Sistema pronto para lhe ajudar", 5, 1)




; Leitura X Host

Local $aMem = MemGetStats()


#cs
	Local $temperatura
	$wbemFlagReturnImmediately = 0x10
	$wbemFlagForwardOnly = 0x20
	$strComputer = "localhost"
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\wmi")
	$Instances = $objWMIService.InstancesOf("MSAcpi_ThermalZoneTemperature")
	For $Item in $Instances
	$temperatura=($Item.CurrentTemperature - 2732) / 10
	Next

	;MsgBox(4096, "name processador ",$temperatura)

#CE
;-------------------------------------------------Processador-----------
$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"

$Cpu_Output = ""
$Cpu_Core = 0
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Processor", "WQL", _
		$wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) Then
	For $objItem In $colItems
		$Cpu_Output = $Cpu_Output & "Architecture: " & $objItem.Architecture & @CRLF
		$Cpu_Output = $Cpu_Output & "Max Clock Speed: " & $objItem.MaxClockSpeed & @CRLF
		$Cpu_Output = $Cpu_Output & "Name: " & $objItem.Name & @CRLF
		$Cpu_Name = $objItem.Name
		$Cpu_Clock = $objItem.MaxClockSpeed
		$Cpu_Arch = $objItem.Architecture
		$Cpu_Core += 1
	Next
Else
	MsgBox(0, "WMI Output", "No WMI Objects Found for class: " & "Win32_Processor")
EndIf

;MsgBox(4096, "name processador ",$Cpu_Name)

;=========== LISTA JANELAS ABERTAS ===================

$oShell = ObjCreate("shell.application") ; Get the Windows Shell Object
$oShellWindows = $oShell.windows ; Get the collection of open shell Windows

If IsObj($oShellWindows) Then

	$janelas = "" ; String for displaying purposes
	For $Window In $oShellWindows ; Count all existing shell windows
		;$janelas = $janelas & $Window.LocationName & Chr(13) & "  |  " & Chr(13) ;<tr><td>Jill</td><td>Smith</td><td>50</td></tr><tr><td>Eve</td><td>Jackson</td><td>94</td></tr><tr><td>John</td><td>Doe</td><td>80</td> </tr>
		$janelas = $janelas & '<tr><td>' & $Window.LocationName & '</td></tr>'
	Next

EndIf

;============= LISTA SERVIÇOS======================

$ObjWMI = ObjGet("winmgmts://" & @ComputerName)
$services = ""
For $item In $ObjWMI.ExecQuery("select * from win32_service")
	;$services = $services & $item.name &"    |    " & @CRLF
	;$services = $services & $item.name & Chr(13) & "  |  " & Chr(13)
	$services = $services & '<tr><td>' & $item.name & '</td></tr>'
Next

;=========== LISTA DRIVERS==============
$objcol = ObjGet("winmgmts:")
$instance = $objcol.instancesof("Win32_LogicalDisk")
If @error Then
	MsgBox(0, "", "error getting object. Error code: " & @error)
	Exit
EndIf

$drivers = ""

For $Drive In $instance




	;$driver_space = int(StringReplace(DriveSpaceFree($Drive.deviceid), ".", "") / 1073741824)
	$driver_space = DriveSpaceFree($Drive.deviceid)
	$driver_space1 = StringSplit($driver_space, ".") ; Split the string of days using the delimiter "," and the default flag value.
	$driver_space = $driver_space1[1]
	$driver_space = StringReplace($driver_space, " ", "")
	$sSerial = DriveGetSerial($Drive.deviceid & "\") ; Find the serial number of the home drive, generally this is the C:\ drive.
	
	If $driver_space = 0 Then
		; MsgBox(4096, "opcao 1 ",$driver_space&" = "&$disk_warning_size&" = "&$disk_critical_size&"    "&DriveSpaceFree($Drive.deviceid)&"  "&$Drive.deviceid& "  "&$Drive.size )

		$img_warning_disk = '<img src="' & $api_web_img & 'alert1.png" alt="' & $Drive.deviceid & '" width="50" height="50">'

	ElseIf $disk_critical_size < $driver_space Then
		
		;MsgBox(4096, "opcao 2 ",$driver_space&" < "&$disk_critical_size&" Critical "&$disk_warning_size&"    "&DriveSpaceFree($Drive.deviceid)&"  "&$Drive.deviceid& "  "&$Drive.size )

		$img_warning_disk = '<img src="' & $api_web_img & 'alert.png" alt="' & $Drive.deviceid & '" width="50" height="50">'
		
	ElseIf $disk_warning_size < $driver_space Then
		
		; MsgBox(4096, "opcao 3 ",$driver_space&" < "&$disk_warning_size&" Warning "&$disk_critical_size&"    "&DriveSpaceFree($Drive.deviceid)&"  "&$Drive.deviceid& "  "&$Drive.size )

		$img_warning_disk = '<img src="' & $api_web_img & 'alert2.png" alt="' & $Drive.deviceid & '" width="50" height="50">'
		
	Else
		
		;  MsgBox(4096, "opcao 4 ",$driver_space&" = "&$disk_warning_size&" = "&$disk_critical_size&"    "&DriveSpaceFree($Drive.deviceid)&"  "&$Drive.deviceid& "  "&$Drive.size )

		$img_warning_disk = '<img src="' & $api_web_img & 'trust.png" alt="' & $Drive.deviceid & '" width="50" height="50">'
		
	EndIf
	

	
	;$str = StringReplace($str, " ", "+")
	;MsgBox(4096, "Lista =",$driver_space&"  "&DriveSpaceFree($Drive.deviceid)&"  "&$Drive.deviceid& "  "&$Drive.size )

	;$drivers = $drivers & $Drive.size & @TAB & $Drive.deviceid & Chr(13) & "  |  " & Chr(13)
	$drivers = $drivers & '<tr><td>' & $img_warning_disk & '</td><td><img src="' & $api_web_img & 'driver.png" alt="' & $Drive.deviceid & '" width="50" height="50">Unidade  ' & $Drive.deviceid & '</td><td><img src="' & $api_web_img & 'size.png" alt="' & $Drive.size & '" width="50" height="50">Disk Size: ' & Int($Drive.size / 1073741824) & ' GB   |   Free space: ' & $driver_space & ' MB | Serial Number: ' & $sSerial & '</td></tr>' & Chr(13) ;  <tr><td>"& $item.name &"</td><td>"& $item.name &"</td><td>"& $item.name &"</td></tr><tr><td>"& $item.name &"</td><td>"& $item.name &"</td><td>"& $item.name &"</td></tr>
	;$drivers = $drivers & $Drive.size & @TAB & $Drive.deviceid & Chr(13) & "  |  " & Chr(13)
	$driver_space = 0
Next

;========= LISTA Process =============

;local $info_file_trust $info_file_win

$ObjWMI = ObjGet("winmgmts://" & @ComputerName)
$process = ""


For $item In $ObjWMI.ExecQuery("select * from Win32_Process")


	;$file_trust $file_windows $file_alert $file_danger
	
	$name = StringReplace($item.name, ".exe", "")
	
	If StringRight($item.ExecutablePath, 4) = ".exe" Then
		$patch_file = $item.ExecutablePath
		$img_path = $api_web_img & "folder.png"
		$patch_alert = $api_web_img & "ok2.png"
		;FileWriteLine(@ScriptDir & "\file_trust.txt", $item.name &";")
	Else
		$patch_file = "Not path folder... service or memory."
		$img_path = $api_web_img & "nofolder.png"
		$patch_alert = $api_web_img & "alert1.png"
		;FileWriteLine(@ScriptDir & "\file_windows.txt", $item.name &";")
	EndIf
	
	;$file_trust $file_windows $file_alert $file_danger
	
	
	
	
	; inicio split file trust
	Local $files_splits2 = StringSplit($file_trust, ";") ; Split the string of days using the delimiter "," and the default flag value.
	; MsgBox(4096, "Lista =",$files_splits2[0])
	
	For $i = 1 To $files_splits2[0] ; Loop through the array returned by StringSplit to display the individual values.
		
		;FileWriteLine(@ScriptDir & "\comparacao.txt", $item.name &"#"&$files_splits2[$i]&"#"&@CRLF)
		If $item.name = $files_splits2[$i] Then
			$info_file_trust = $api_web_img & "trust.png"
			;MsgBox(4096, "Lista =",$item.name &"  "&$files_splits2[$i])
		EndIf

	Next
	; fim split file trust
	
	; inicio split file windows
	Local $files_splits2 = StringSplit($file_windows, ";") ; Split the string of days using the delimiter "," and the default flag value.
	; MsgBox(4096, "Lista =",$files_splits2[0])
	
	For $i = 1 To $files_splits2[0] ; Loop through the array returned by StringSplit to display the individual values.
		
		;FileWriteLine(@ScriptDir & "\comparacao.txt", $item.name &"#"&$files_splits2[$i]&"#"&@CRLF)
		If $item.name = $files_splits2[$i] Then
			
			$info_file_trust = $api_web_img & "trust.png"
			$info_file_win = '<img src="' & $api_web_img & 'windows.png" alt="' & $item.name & '" width="50" height="50">'
			;MsgBox(4096, "Lista =",$item.name &"  "&$files_splits2[$i])
		EndIf

	Next
	; fim split file trust
	
	
	
	; inicio split file alert
	Local $files_splits2 = StringSplit($file_alert, ";") ; Split the string of days using the delimiter "," and the default flag value.
	; MsgBox(4096, "Lista =",$files_splits2[0])
	
	For $i = 1 To $files_splits2[0] ; Loop through the array returned by StringSplit to display the individual values.
		
		;FileWriteLine(@ScriptDir & "\comparacao.txt", $item.name &"#"&$files_splits2[$i]&"#"&@CRLF)
		If $item.name = $files_splits2[$i] Then
			
			$info_file_trust = $api_web_img & "ok.png"
			;MsgBox(4096, "Lista =",$item.name &"  "&$files_splits2[$i])
		EndIf

	Next
	; fim split file alert
	
	
	
	; inicio split file danger
	Local $files_splits2 = StringSplit($file_danger, ";") ; Split the string of days using the delimiter "," and the default flag value.
	; MsgBox(4096, "Lista =",$files_splits2[0])
	
	For $i = 1 To $files_splits2[0] ; Loop through the array returned by StringSplit to display the individual values.
		
		;FileWriteLine(@ScriptDir & "\comparacao.txt", $item.name &"#"&$files_splits2[$i]&"#"&@CRLF)
		If $item.name = $files_splits2[$i] Then
			
			$info_file_trust = $api_web_img & "danger.png"
			;MsgBox(4096, "Lista =",$item.name &"  "&$files_splits2[$i])
		EndIf

	Next
	; fim split file danger
	
	; <a href="http://www.w3schools.com/" target="_blank">Visit W3Schools!</a>
	;$process = $process & $item.name &"  #"&@CRLF & Chr(13) & @lf
	;$process = $process & $item.name & "  |  " & $item.ExecutablePath & Chr(13) & "  |  " & Chr(13) <img src="w3schools.jpg" alt="W3Schools.com" width="104" height="142">
	$process = $process & '<tr><td><img src="' & $info_file_trust & '" alt="' & $item.name & '" width="50" height="50"></td><td>' & $info_file_win & '<img src="' & $api_web_img & $name & '.png" alt="' & $item.name & '" width="50" height="50">' & $item.name & '</td><td><img src="' & $img_path & '" alt="' & $item.name & '" width="50" height="50">' & $patch_file & '</td><td><img src="' & $patch_alert & '" alt="' & $item.name & '" width="50" height="50"></td><td><a href="http://searchtasks.answersthatwork.com/tasklist.php?File=' & $name & '" target="_blank"><img src="' & $api_web_img & 'info.png" alt="' & $item.name & '" width="50" height="50"></a></td></tr>' & Chr(13) ;  <tr><td>"& $item.name &"</td><td>"& $item.name &"</td><td>"& $item.name &"</td></tr><tr><td>"& $item.name &"</td><td>"& $item.name &"</td><td>"& $item.name &"</td></tr>
	;$process = $process & $item.name &" | "& $item.ExecutablePath&" | "&$item.ParentProcessId &" | "& $item.GetOwner&" | "&$item.Priority &" | "& $item.WorkingSet&" | "&$item.PercentProcessorTime & @CRLF
	
	$info_file_trust = $api_web_img & "alert.png"
	$info_file_win = ""
	#cs

		<tr><th>Lista</th><th>Processo</th><th>Caminho</th><th>estrutura</th><th>Informações</th></tr>


		<tr><td>Jill</td><td>Smith</td><td>50</td></tr><tr><td>Eve</td><td>Jackson</td><td>94</td></tr><tr><td>John</td><td>Doe</td><td>80</td> </tr>


	#ce
	
Next



;Loop esperando a interação do usuário
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg

		;_About($about, $MainLabel, $CopyRLabel, StringLeft(FileGetVersion(@ScriptName), 5), $NameURL1, $URL1, $NameURL2, $URL2, $NameURL3, $URL3, @ScriptName, $LinkColor, $BkColor, -1, -1, -1, -1, $hWnd)

		Case $curtir_face

			#cs
				; Use no proxy
				HttpSetProxy(1)

				; Use IE defaults for proxy
				HttpSetProxy(0)
				; Use the proxy "www-cache.myisp.net" on port 8080
				HttpSetProxy(2, "www-cache.myisp.net:8080")
			#ce
			ShellExecute("http://www.facebook.com/pages/Bhns/430524047029097", @SW_MAXIMIZE, "", "open")
			TrayTip("AIT HelpDesk ®", "Obrigado " & $inputNome & @LF & " ficamos satisfeitos...." & @LF & " AIT HelpDesk ® agradece.", 5, 1)
			

		Case $GUI_EVENT_CLOSE
			Local $iAnswer = MsgBox(36, "Pergunta", "Sair sem enviar solicitação?")
			;If iAnswer = Sim
			If $iAnswer = 6 Then
				;MsgBox(64, "Ação cancelada", "Nenhuma alteração realizada.")
				DirRemove($last_tickets & $random_n, 1)
				;Traduz os inputs de int para String
				$inputMensagem = GUICtrlRead($formMensagem)
				$inputNome = GUICtrlRead($formNome)
				$inputEmail = GUICtrlRead($formEmail)
				$inputBreveDescr = GUICtrlRead($formBreveDescr)
				$inputDept = GUICtrlRead($formDept)
				$inputPriorit = GUICtrlRead($tipo)
				$inputtelefone = GUICtrlRead($input_telefone)
				$inputRamal = GUICtrlRead($Input_ramal)
				$helptopic = $inputDept
				GUIDelete()
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "USUARIO", $inputNome)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EMAIL", $inputEmail)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "PHONE", $inputtelefone)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EXTENSION", $inputRamal)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_ASSUNTO", $inputBreveDescr)
				_sair()
				TrayTip("AIT HelpDesk ®", " Agradeçemos sua preferência, obrigado equipe AIT Helpdesk...", 5, 1)

				
				ExitLoop
				Exit
			EndIf

			;Botão Cancelar
		Case $botaoCancelar
			Local $iAnswer = MsgBox(36, "Pergunta", "Sair sem enviar solicitação?")
			;If iAnswer = Sim
			If $iAnswer = 6 Then
				;MsgBox(64, "Ação cancelada", "Nenhuma alteração realizada.")
				DirRemove($last_tickets & $random_n, 1)
				;Traduz os inputs de int para String
				$inputMensagem = GUICtrlRead($formMensagem)
				$inputNome = GUICtrlRead($formNome)
				$inputEmail = GUICtrlRead($formEmail)
				$inputBreveDescr = GUICtrlRead($formBreveDescr)
				$inputDept = GUICtrlRead($formDept)
				$inputPriorit = GUICtrlRead($tipo)
				$inputtelefone = GUICtrlRead($input_telefone)
				$inputRamal = GUICtrlRead($Input_ramal)
				$helptopic = $inputDept
				GUIDelete()
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "USUARIO", $inputNome)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EMAIL", $inputEmail)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "PHONE", $inputtelefone)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EXTENSION", $inputRamal)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_ASSUNTO", $inputBreveDescr)
				_sair()
				TrayTip("AIT HelpDesk ®", " Agradeçemos sua preferência, obrigado equipe AIT Helpdesk...", 5, 1)
				;TrayTip("Equipe AIT HelpDesk ®", " Agradeçe a você "&$inputNome&"...", 5, 1)
				

				ExitLoop
				Exit
			EndIf
			;Botão Enviar
		Case $botaoEnviar




			;Traduz os inputs de int para String
			$inputMensagem = GUICtrlRead($formMensagem)
			If $inputMensagem = '' Then
				$inputMensagem = 'SEM ASSUNTO'
			EndIf
			
			$inputNome = GUICtrlRead($formNome)
			$inputEmail = GUICtrlRead($formEmail)
			$inputBreveDescr = GUICtrlRead($formBreveDescr)
			If $inputBreveDescr = '' Then
				$inputBreveDescr = 'SEM DESCRICAO'
			EndIf
			
			
			$inputDept = GUICtrlRead($formDept)
			$inputPriorit = GUICtrlRead($tipo)
			$inputtelefone = GUICtrlRead($input_telefone)
			$inputRamal = GUICtrlRead($Input_ramal)
			$helptopic = $inputDept
			GUIDelete()





			;Transforma o departamento em int
			If $inputDept = $name_DPN_1 Then
				$departamento = $id_DPN_1
				$topic_ID_ = $tpid_DPN_1
				$flagAnexo = 0

			ElseIf $inputDept = $name_DPN_2 Then
				$departamento = $id_DPN_2
				$topic_ID_ = $tpid_DPN_2
				$flagAnexo = 1

			ElseIf $inputDept = $name_DPN_3 Then
				$departamento = $id_DPN_3
				$topic_ID_ = $tpid_DPN_3
				$flagAnexo = 1

			ElseIf $inputDept = $name_DPN_4 Then
				$departamento = $id_DPN_4
				$topic_ID_ = $tpid_DPN_4
				$flagAnexo = 1


			ElseIf $inputDept = $name_DPN_5 Then
				$departamento = $id_DPN_5
				$topic_ID_ = $tpid_DPN_5
				$flagAnexo = 1

			ElseIf $inputDept = $name_DPN_6 Then
				$departamento = $id_DPN_6
				$topic_ID_ = $tpid_DPN_6
				$flagAnexo = 1

			ElseIf $inputDept = $name_DPN_7 Then
				$departamento = $id_DPN_7
				$topic_ID_ = $tpid_DPN_7
				$flagAnexo = 1
			Else
				$departamento = 1
				$topic_ID_ = $tpid_DPN_1
				
			EndIf


			; ============TICKET COMUM =======================
			If $inputPriorit = $prioridade_id1 Then
				$priorit_id = 2
				$status_remot = 1
				$RDP = ""

			ElseIf $inputPriorit = $prioridade_id2 Then
				$priorit_id = 3
				$status_remot = 1
				$RDP = ""
				;=================VNC HABILITADO =====================
			ElseIf $inputPriorit = $prioridade_id3 Then
				$priorit_id = 3
				$status_remot = 3
				$empresa = $empresa & " | " & $inputPriorit
				criar_profile_vpn()
				enabled_vnc_assitencie() ; HABILITA VNC
				Sleep(3000)

				;=================RDP HABILITADO =====================
			ElseIf $inputPriorit = $prioridade_id4 Then
				$priorit_id = 4
				$status_remot = 2
				$empresa = $empresa & " | " & $inputPriorit
				criar_profile_vpn()
				open_ports_rdp()
				enabled_rdp_assitencie() ; HABILITA ASSISTENCIA REMOTA
				Sleep(3000)

			Else
				$RDP = " | " & $inputPriorit
				$priorit_id = 2
				$status_remot = 1

			EndIf


			
			
			
			
			
			; Gerar relatorio por grupo informado
			
			Local $report_file_html = ""
			Local $report_group = StringSplit($report_workstation, ";") ; Split the string of days using the delimiter "," and the default flag value.
			; MsgBox(4096, "Lista =",$report_group[0])
			
			For $i = 1 To $report_group[0] ; Loop through the array returned by StringSplit to display the individual values.
				
				;FileWriteLine(@ScriptDir & "\comparacao.txt", $item.name &"#"&$report_group[$i]&"#"&@CRLF)
				If $departamento = $report_group[$i] Then
					
					
					$memory_load = $aMem[0] ; Memory Load (Percentage of memory in use)
					$memory_total = $aMem[1] ; Total physical RAM (KB)
					$memory_available = $aMem[2] ; Available physical RAM
					$memory_tpage = $aMem[3] ; Total Pagefile
					$memory_avaPage = $aMem[4] ; Available Pagefile
					$memory_tVT = $aMem[5] ; Total virtual
					$memory_aVT = $aMem[6] ; Available virtual
					;<tr><td><img src="' & $api_web_img & 'memory.png" alt="' & $memory_total & '" width="50" height="50">'$memory_total&' KB em uso '$memory_load&'% Livre '$memory_available&'KB</td></tr>



					$dados_user_form = '<tr><td><img src="' & $api_web_img & 'user1.png" alt="' & $usuarioLogado & '" width="50" height="50"><img src="' & $api_web_img & 'user.png" alt="' & $usuarioLogado & '" width="50" height="50">' & $usuarioLogado & '</td></tr><tr><td><img src="' & $api_web_img & 'computer.png" alt="' & $nomeComputador & '" width="50" height="50">' & $nomeComputador & '</td></tr><tr><td><img src="' & $api_web_img & 'processor.png" alt="' & $nomeComputador & '" width="50" height="50">' & $Cpu_Name & '  ' & @CPUArch & '</td></tr><tr><td><img src="' & $api_web_img & 'memory.png" alt="' & $memory_total & '" width="50" height="50">' & Int($memory_total / 1024) & ' MB | Uso ' & $memory_load & '% | Livre ' & Int($memory_available / 1024) & ' MB</td></tr><tr><td><img src="' & $api_web_img & 'sistema.png" alt="' & $nomeComputador & '" width="50" height="50">' & @OSVersion & ' ' & @OSArch & '  ' & $servicePack & '</td></tr><tr><td><img src="' & $api_web_img & 'domain.png" alt="' & $nomeComputador & '" width="50" height="50">' & @LogonDNSDomain & "  " & @LogonDomain & "  " & @LogonServer & '</td></tr><tr><td><img src="' & $api_web_img & 'network.png" alt="' & $usuarioLogado & '" width="50" height="50"> Adapter 1 = ' & @IPAddress1 & '</td></tr><tr><td><img src="' & $api_web_img & 'network.png" alt="' & $usuarioLogado & '" width="50" height="50">  Adapter 2 = ' & @IPAddress2 & '</td></tr><tr><td><img src="' & $api_web_img & 'network.png" alt="' & $usuarioLogado & '" width="50" height="50"> Adapter 3 = ' & @IPAddress3 & '</td></tr><tr><td><img src="' & $api_web_img & 'network.png" alt="' & $usuarioLogado & '" width="50" height="50">  Adapter 4 = ' & @IPAddress4 & '</td></tr>'

					#cs

						@HomeShare
						<img src="' & $api_web_img & 'info2.png" alt="' & $usuarioLogado & '" width="50" height="50">
						<img src="' & $api_web_img & 'storage.png" alt="' & $usuarioLogado & '" width="50" height="50">
						<img src="' & $api_web_img & 'process.png" alt="' & $usuarioLogado & '" width="50" height="50">
						<img src="' & $api_web_img & 'window.png" alt="' & $usuarioLogado & '" width="50" height="50">

						</tr><tr>
						<tr></tr>
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>Usuario Local: " & $usuarioLogado & "</p>")  $Cpu_Name
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>Computador: " & $nomeComputador & "</p>")
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>Sistema: " & $sistemaOperacional & " " & $servicePack & "</p>")
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>User Local: " & $usuarioLogado & " " & $servicePack & "</p>")
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>Domains & Servers: " & @LogonDNSDomain & ", " & @LogonDomain & ", " & @LogonServer & "</p>")
						FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>Endereços de IP: " & @IPAddress1 & ", " & @IPAddress2 & ", " & @IPAddress3 & ", " & @IPAddress4 & "</p>")
					#ce
					;Gerar HTML Formatado

					Local $html_file = "REPORT-WORKSTATION-" & $nomeComputador & ".html"
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<!DOCTYPE html>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<html>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<head>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<style>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "#header {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "background-color:#A6CAF0;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "color:black;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "text-align:center;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "padding:1px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "}")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "#nav {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "line-height:30px;")
					;FileWriteLine($last_tiket_reports & "\" & $html_file, "background-color:#eeeeee;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "height:100%;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "width:300px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "float:left;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "padding:5px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "}")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "#section {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "width:1000px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "float:left;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "padding:10px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "}")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "#footer {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "background-color:#A6CAF0;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "color:white;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "clear:both;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "text-align:center;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "padding:5px;	")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "}")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "table, th, td {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "border: 1px solid black;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "padding: 5px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "}")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "table {")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "border-spacing: 1px;")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</style>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</head>") ;
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<body>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<div id="header">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<img src="' & $api_web_img & 'ait.png" alt="' & $usuarioLogado & '" width="150" height="150">')
					;FileWriteLine($last_tiket_reports & "\" & $html_file, '</br>')
					;FileWriteLine($last_tiket_reports & "\" & $html_file, '<h2>Usuario Local: '& $usuarioLogado &'     Computador: '& $nomeComputador&'</h2>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '</div>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<div id="nav">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<table style="width:100%">')
					;FileWriteLine($last_tiket_reports & "\" & $html_file, '<img src="' & $api_web_img & 'icone.png" alt="' & $usuarioLogado & '" width="100" height="100">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<h1><img src="' & $api_web_img & 'info2.png" alt="' & $usuarioLogado & '" width="25" height="25">    Informações</h1>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, $dados_user_form)
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</table>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<table style="width:100%">')
					;FileWriteLine($last_tiket_reports & "\" & $html_file, '')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<h1><img src="' & $api_web_img & 'services.png" alt="' & $usuarioLogado & '" width="25" height="25">    Serviços</h1>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>" & $services & "</p>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</table>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '</div>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<div id="section">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<h1><img src="' & $api_web_img & 'storage.png" alt="' & $usuarioLogado & '" width="25" height="25">    Armazenamento</h1>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<table style="width:100%">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, $drivers)
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</table>")
					;FileWriteLine($last_tiket_reports & "\" & $html_file, "<p></p>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<h1><img src="' & $api_web_img & 'process.png" alt="' & $usuarioLogado & '" width="25" height="25">    Processos em execução</h1>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<table style="width:100%">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, $process)
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</table>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<h1><img src="' & $api_web_img & 'window.png" alt="' & $usuarioLogado & '" width="25" height="25">    Janelas Ativas</h1>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<p></p>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<table style="width:100%">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<p>" & $janelas & "</p>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</table>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "<p></p>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '</div>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, '<div id="footer">')
					FileWriteLine($last_tiket_reports & "\" & $html_file, "Copyright © bhns")
					FileWriteLine($last_tiket_reports & "\" & $html_file, '</div>')
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</body>")
					FileWriteLine($last_tiket_reports & "\" & $html_file, "</html>")

					Local $sFilePath = $last_tiket_reports & "\" & $html_file

					; Create a temporary file to read data from.
					If Not FileCreate($sFilePath, "This is an example of using FileRead.") Then Return MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")

					; Open the file for reading and store the handle to a variable.
					Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
					If $hFileOpen = -1 Then
						MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
						Return False
					EndIf

					; Read the contents of the file using the handle returned by FileOpen.
					Local $sFileRead = FileRead($hFileOpen)

					; Close the handle returned by FileOpen.
					FileClose($hFileOpen)

					; Display the contents of the file.
					;MsgBox($MB_SYSTEMMODAL, "", $sFileRead)

					; Delete the temporary file.
					;FileDelete($sFilePath)


					$file_encode64 = _Base64encode($sFileRead, 0)

					;FileWrite("E:\FORMULARIO\teste.txt", $file_encode64)

					$report_file_html = '{"' & $html_file & '":"data:text/html;base64,' & $file_encode64 & '"}'

					Local $report_station = '<h1>AIT HelpDesk Reports</h1></br><p></p><h2>Informacoes da maquina cliente</h4><p>Usuario Local: ' & $usuarioLogado & '</p><p>Computador: ' & $nomeComputador & '</p><p>Sistema: ' & $sistemaOperacional & ' ' & $servicePack & '</p><p>User Local: ' & $usuarioLogado & ' ' & $servicePack & '</p><p>Domains & Servers: ' & @LogonDNSDomain & ', ' & @LogonDomain & ', ' & @LogonServer & '</p><p></p><h2>Networks</h4><p></p><p>Enderecos de IP: ' & @IPAddress1 & ', ' & @IPAddress2 & ', ' & @IPAddress3 & ', ' & @IPAddress4 & '</p><p></p><h2>JANELAS ATIVAS</h4><p></p><p></p><p></p><h2>Processos</h4><p></p><p></p><p></p><h2>Servicos</h4><p></p><p></p><p></p><h2>Fim Reports...</h4>'

					
					
				Else
					
					
					
				EndIf

			Next
			; fim split file trust
			
			
			;MsgBox($MB_SYSTEMMODAL, "", $report_file_html)
			
			
			
			
			
			
			
			;====================Copia os prints e reports para o servidor============
			;MsgBox(64, "pasta imagens local", $tiket_img_dir)
			;MsgBox(64, "pasta imagens externo", $dir_tickets)
			;MsgBox(64, "pasta reports local", $last_tiket_reports & "\" & $html_file)
			;MsgBox(64, "pasta reports externo", $dir_tickets)

			TrayTip("AIT HelpDesk ®", " Registrando Ticket....," & @LF & " Aguarde...", 5, 1)



			
			

			;=======================ANEXOS TICKETS===========================

			TrayTip("AIT HelpDesk ®", " Enviando anexos." & @LF & " Aguarde.....", 5, 1)

			If $aPrtScreenSize > 1 Then

				; Envia Anexos
				For $i = 0 To ($aPrtScreenSize - 2)
					$file_ = $file_ & $arrayPrtScreen[$i]
					If $i <> ($aPrtScreenSize - 2) Then
						$file_ = $file_ & ";"

					EndIf
				Next

				;MsgBox(4096, "array",$file_ )
				;MsgBox(4096, "array quebra de linha",$file_& @CRLF )

				Local $S_files_ = StringSplit($file_, ";")
				For $x = 1 To $S_files_[0]
					$S_files_[$x] = _PathFull($S_files_[$x])
					If FileExists($S_files_[$x]) Then
						;ConsoleWrite('+> File attachment added: ' & $S_files_[$x] & @LF)

					Else
						;ConsoleWrite('!> File not found to attach: ' & $S_files_[$x] & @LF)
						SetError(1)
						Return 0
					EndIf

					;MsgBox(4096, "array sequencial ",$S_files_[$x])
					Local $namefile_ = $S_files_[$x]
					Local $file_full = StringRight($S_files_[$x], 30)
					Local $file_name_ = StringRight($S_files_[$x], 10)
					Local $file_key_ = StringLeft($file_full, 15)
					;MsgBox(4096, "Diretorio",$diretorio)
					Local $get_file_size = FileGetSize($S_files_[$x])
					;MsgBox(4096, "Arquivo direita",$file_full)
					;MsgBox(4096, "Arquivo direita 2",$file_name_)
					;MsgBox(4096, "SIZE",$get_file_size)
					;MsgBox(4096, "Loop = ",$aPrtScreenSize)
					
					

					Local $files_splits2 = StringSplit($namefile_, "\") ; Split the string of days using the delimiter "," and the default flag value.

					For $i = 1 To $files_splits2[0] ; Loop through the array returned by StringSplit to display the individual values.
						
						If $i = 1 Then
							$file_name_ = $files_splits2[$i]
							;MsgBox(4096, "Diretorio =",$file_name_)
							
						ElseIf $i >= 2 Then
							$file_name_ = $files_splits2[$i]
							;MsgBox(4096, "Arquivo =",$file_name_)
						Else
							
						EndIf

					Next
					
					
					
					
					

					; Create a constant variable in Local scope of the filepath that will be read/written to.
					Local $sFilePath = $namefile_

					; Create a temporary file to read data from.
					If Not FileCreate($sFilePath, "This is an example of using FileRead.") Then Return MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")

					; Open the file for reading and store the handle to a variable.
					Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
					If $hFileOpen = -1 Then
						MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file: " & $sFilePath)
						Return False
					EndIf

					; Read the contents of the file using the handle returned by FileOpen.
					Local $sFileRead = FileRead($hFileOpen)

					; Close the handle returned by FileOpen.
					FileClose($hFileOpen)

					; Display the contents of the file.
					;MsgBox($MB_SYSTEMMODAL, "", $sFileRead)

					; Delete the temporary file.
					;FileDelete($sFilePath)
					; Create a file.
					Func FileCreate($sFilePath, $sString)
						Local $bReturn = True ; Create a variable to store a boolean value.
						If FileExists($sFilePath) = 0 Then $bReturn = FileWrite($sFilePath, $sString) = 1 ; If FileWrite returned 1 this will be True otherwise False.
						Return $bReturn ; Return the boolean value of either True of False, depending on the return value of FileWrite.
					EndFunc   ;==>FileCreate
					$file_encode64 = _Base64encode($sFileRead, 0)
					;FileWrite("E:\FORMULARIO\teste.txt", $file_encode64)
					$json_attachments = '{"' & urlencode($file_name_) & '":"data:image/jpeg;base64,' & $file_encode64 & '"}'
					; MsgBox(4096, "teste ","tamanho "& $x &"outro")
					If $x = 1 Then

						$json_attachments_full = $json_attachments

					ElseIf $x > 1 Then
						
						$json_attachments_full = $json_attachments_full & ',' & $json_attachments
						
					EndIf
					
					
				Next
				
				
				
				
				; Solucao para enviar relatorio somente para o grupo selecionado no final dos anexos
				If $report_file_html = "" Then
				Else
					$json_attachments_full = $json_attachments_full & ',' & $report_file_html
				EndIf
				;======================================================================
				
				
				
				



			Else
				
				$json_attachments_full = $report_file_html
				
				
				
				
				
				
				

			EndIf
			
			
			
			
			
			
							;=========================================Enviar dados de acesso remoto========================================
				
				
				If $status_remot = 3 Then
					
					;$dados_host = "AIT HelpDesk - Reports </br> Report link interno: " & $Link_anexo_interno & $html_file & " </br> Report link Externo: " & $Link_anexo_externo & $html_file & " </br> AIT HelpDesk - Assistencia VNC habilitada </br> Horario de inicio : " & $horario & " </br> Password : 70x7  </br> IPs de acessos liberados" & @IPAddress1 & ":6064" & ", " & @IPAddress2 & ":6064" & ", " & @IPAddress3 & ":6064" & ", " & @IPAddress4 & ":6064" & " </br> Por padrao um IP e da rede local do usuario, um e fornecido pela VPN que e o mais utilizado caso o cliente esteje remoto e o restante caso o cliente tenha mais de uma conexao </br> Equipe Ait HelpDesk </br> Perdoar e preciso sempre...."
					
					;$info_remot = $anexo_html
					$key_rdp = "AIT HelpDesk - Assistencia VNC habilitada "
					Local $file_name_ = "VNC-ID-" & $random_n & ".vnc"
					Local $file_key_ = "AITVNC" & @YEAR & @MON & $random_rdp
					;MsgBox(4096, "Diretorio",$diretorio)
					Local $get_file_size = FileGetSize($folder_convicts & $file_rdp)
					;MsgBox(4096, "Diretorio",$folder_convicts & $file_rdp)
					
				ElseIf $status_remot = 2 Then
              
					Local $file_anexo_name = "RDP_PASSWD_AITKEY" & $random_KEY&".msrcincident"
					Local $extension_file = "msrcincident"
				    Local $file_anexo
					anexo_file()
					
					If $json_attachments_full = "" then
					
					$json_attachments_full = $file_anexo
					
					else
					
					$json_attachments_full = $json_attachments_full & ',' & $file_anexo
					
					
					EndIf
					
				Else
					;$dados_host = "AIT HelpDesk - Reports </br> Report link interno: " & $Link_anexo_interno & $html_file & " </br> Report link Externo: "& $Link_anexo_externo & $html_file
				EndIf
				;=============================================================================================================
				
			;
			;FileWrite("E:\FORMULARIO\json_parse.txt", $json_attachments_full)

			; Envio de email

			If $aPrtScreenSize - 2 < 0 Then
				TrayTip("Aviso", "A solicitação está sendo enviada." & @LF & "Por favor, aguarde.", 5, 1)
			ElseIf $aPrtScreenSize - 2 = 0 Then
				TrayTip("Aviso", "A solicitação está sendo enviada com " & $aPrtScreenSize - 1 & " anexo." & @LF & "Por favor, aguarde.", 5, 1)
			Else
				TrayTip("Aviso", "A solicitação está sendo enviada com " & $aPrtScreenSize - 1 & " anexos." & @LF & "Por favor, aguarde.", 5, 1)
			EndIf


			If $aPrtScreenSize > 1 Then
				;MsgBox(4096, "Caminho arquivo",$dir_tickets)
				

			Else

				$caminho = " Sem anexos enviados"

			EndIf

			;#cs
			
			
			;; API OSTICKET INICIO


			#CS

				;Traduz os inputs de int para String
				$inputMensagem = GUICtrlRead($formMensagem)
				$inputNome = GUICtrlRead($formNome)
				$inputEmail = GUICtrlRead($formEmail)
				$inputBreveDescr = GUICtrlRead($formBreveDescr)
				$inputDept = GUICtrlRead($formDept)
				$inputPriorit = GUICtrlRead($tipo)
				$inputtelefone = GUICtrlRead($input_telefone)
				$inputRamal = GUICtrlRead($Input_ramal)
				$helptopic = $inputDept

			#CE
			

			
			
			#cs   SAMPLE JSON

                https://github.com/RockefellerArchiveCenter/osTicket/blob/master/setup/doc/api/tickets.md

				{
				"alert": true,
				"autorespond": true,
				"source": "API",
				"name": "Angry User",
				"email": "api@osticket.com",
				"phone": "3185558634X123",
				"subject": "Testing API",
				"ip": "123.211.233.122",
				"message": "MESSAGE HERE",
				"attachments": [
				{"file.txt": "data:text/plain;charset=utf-8,content"},
				{"image.png": "data:image/png;base64,R0lGODdhMAA..."},
				]
				}


			#ce
			
			Global $sdata_json = '{"alert":' & $report_api_mail & ',"autorespond":true,"source":"API","name":"' & $inputNome & '","email":"' & $inputEmail & '","phone":"' & $inputtelefone & ' ' & $inputRamal & '","subject":"' & urlencode($inputBreveDescr) & '   #' & $empresa & '","message":"data:text/html;charset=utf-8,' & urlencode($inputMensagem) & '","ip":"' & $ipAddress & '","priority":"' & $priorit_id & '","topicId":"' & $topic_ID_ & '","attachments":[ ' & $json_attachments_full & ']}'

			
			
			#CS  FORCE SEND IP  CHANGE FILE CLASS.API.PHP LINE 175 USED $_SERVER['HTTP_USER_AGENT'] FOR SEND IP API


				function requireApiKey() {
				# Validate the API key -- required to be sent via the X-API-Key
				# header



				if(!($key=$this->getApiKey()))
				return $this->exerr(401, __('Valid API key required'));
				elseif (!$key->isActive() || $key->getIPAddr()!=$_SERVER['HTTP_USER_AGENT']) //$_SERVER['REMOTE_ADDR'] $_SERVER['HTTP_USER_AGENT'] or '189.32.32.1' BY cleber
				return $this->exerr(401, __('API key not found/active or source IP not authorized'));

				return $key;


				}

				function getApiKey() {


				if (!$this->apikey && isset($_SERVER['HTTP_X_API_KEY']) && isset($_SERVER['REMOTE_ADDR']) )
				$this->apikey = API::lookupByKey($_SERVER['HTTP_X_API_KEY'], $_SERVER['HTTP_USER_AGENT']);//$_SERVER['REMOTE_ADDR']  , $_SERVER['REMOTE_ADDR'] or '189.32.32.1' BY cleber

				return $this->apikey;

				}

			#CE

			
			;$api_ip
			;MsgBox(8192, "data  ", $sdata_json)
			;MsgBox(8192, "lINK ", $api_host&$api_link)

			$LocalIP = $api_host

			$hw_open = _WinHttpOpen()

			$hw_connect = _WinHttpConnect($hw_open, $LocalIP)

			$h_openRequest = _WinHttpOpenRequest($hw_connect, "POST", $api_link)

			$headers = 'X-API-Key:' & $api_key & ' ' & @CRLF & _
					'User-Agent:' & $api_ip & @CRLF & _
					'Content-Type: text/html;chart=UTF-8' & @CRLF
			
			;MsgBox(8192, "Cabecalho  ", $headers)
			#cs 'Connection: keep-alive' & @CRLF & _
					'Host: '&$api_ip& @CRLF & _
					'REMOTE_ADDR:'&$api_ip&@CRLF & _
					'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' & @CRLF & _
					'Cookie: schemarepo=TN_MP3; userDb=D01; cqecytusid=Zjnh1Smbfl/MiWNLFGixNg==; JSESSIONID=0000F3DyGHdvfvfTMFHI_WA6q2I:-1' & @CRLF & _
					'Referer: http://localhost/osticket/api/tickets.json' & @CRLF & _
					'Accept-Language: pl,en-us;q=0.7,en;q=0.3'& @CRLF
			#ce

			;MsgBox(8192, "Encode64 data", $sdata_json)
			_WinHttpSendRequest($h_openRequest, $headers, $sdata_json)


			_WinHttpReceiveResponse($h_openRequest)
			
			
			Local $response_headers = _WinHttpQueryHeaders($h_openRequest)
			;MsgBox(8192, "Headers", $response_headers)
			Local $headers_splits = StringSplit($response_headers, @CRLF) ; Split the string of days using the delimiter "," and the default flag value.
			Local $headers_status = $headers_splits[1]
			;MsgBox(4096, "Status",$headers_status)
			Local $response1 = _WinHttpReadData($h_openRequest)
			While @extended = 8192
				$response1 &= _WinHttpReadData($h_openRequest)
			WEnd
			

			
			;ConsoleWrite(@extended & @CRLF)
			
			;MsgBox(8192, "Headers", $headers_status)
			FileWriteLine(@ScriptDir & "\log_erro.txt", $headers_status)
			Local $response = $response1
			If $headers_status = "HTTP/1.1 201 Created" Then
				
				retorno_ticket()
			ElseIf $headers_status = "HTTP/1.1 200 OK" Then
				
				$response = StringRight($response, 6)
				retorno_ticket()
			ElseIf $headers_status = "HTTP/1.1 400 Bad Request" Then ;HTTP/1.1 401 Unauthorized
				retorno_ticket_erro()
				FileWriteLine(@ScriptDir & "\log_erro.txt", $headers_status & @CRLF & $response & @CRLF)
				
			ElseIf $headers_status = "HTTP/1.1 401 Unauthorized" Then ;HTTP/1.1 401 Unauthorized
				retorno_ticket_erro()
				FileWriteLine(@ScriptDir & "\log_erro.txt", $headers_status & @CRLF & $response & @CRLF)
				
			Else
				retorno_ticket_erro()
				$headers_status = "Sem conectividade com o servidor" & @CRLF & $api_host
				FileWriteLine(@ScriptDir & "\log_erro.txt", $date_sql & " Sem conectividade com o servidor " & $api_host & " " & $response & @CRLF)
				
			EndIf
			
			
			Func retorno_ticket()
				
				Dim $iMsgBoxAnswer
				$iMsgBoxAnswer = MsgBox(8260, "Servidor HelpDesk  |  " & $LocalIP, "Segue a ID do seu Ticket #" & $response & @CRLF & Chr(13) & "Mais detalhes sera enviado em seu email " & $inputEmail & @CRLF & Chr(13) & "Seu ticket foi registrado com sucesso.." & @CRLF & Chr(13) & "Dejesa vizualizar seu ticket via web ?")
				Select
					Case $iMsgBoxAnswer = 6 ;Yes
						
						ShellExecute("http://" & $api_web_front & "view.php?e=" & $inputEmail & "&t=" & $response, @SW_MAXIMIZE, "", "open") ;http://localhost/osticket/view.php?e=clebe@live.com&t=201504
						TrayTip("AIT HelpDesk ®", "Obrigado " & $inputNome & @LF & "ficamos satisfeitos...." & @LF & "AIT HelpDesk ® agradece.", 5, 1)

					Case $iMsgBoxAnswer = 7 ;No

				EndSelect
				;MsgBox(8192, "Resposta do Servidor  " & $LocalIP, "Segue dados do seu Ticket N: " & $response & @CRLF & Chr(13) & " Mais detalhes sera enviado em seu email " & $inputEmail & @CRLF & Chr(13) & " Seu ticket foi enviado com sucesso..")
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "USUARIO", $inputNome)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EMAIL", $inputEmail)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "PHONE", $inputtelefone)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "EXTENSION", $inputRamal)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_TICKET", $response)
				IniWrite($folder_aithelpdesk & "\profile.ini", "USER", "LAST_ASSUNTO", $inputBreveDescr)
				
				;==============================================sendmail==============================================
				
				
				If $report_user_mail = 0 Then

					enviar_mail()
					
				Else
					
					
				EndIf

				
				
				
				
				
				
				
				
				
			EndFunc   ;==>retorno_ticket

			Func retorno_ticket_erro()

				Dim $iMsgBoxAnswer
				
				$iMsgBoxAnswer = MsgBox(8244, "Erro ao processar Ticket", "Erros encontrados" & @CRLF & Chr(13) & $headers_status & @CRLF & $response & @CRLF & Chr(13) & "Por Favor, Tente novamente ")
				Select
					Case $iMsgBoxAnswer = 6 ;Yes
						
					Case $iMsgBoxAnswer = 7 ;No

				EndSelect
			EndFunc   ;==>retorno_ticket_erro

			_WinHttpCloseHandle($h_openRequest)

			_WinHttpCloseHandle($hw_connect)

			_WinHttpCloseHandle($hw_open)


			;API OS TICKET FIM
			
			

			TrayTip("AIT HelpDesk ®", " Fechando conexão com o banco " & @LF & " Aguarde.........", 5, 1)
			

			TrayTip("AIT HelpDesk ®", "Solicitação enviada com sucesso TicketID #" & $random_n, 5, 1)
			TrayTip("AIT HelpDesk ®", "Para acompanhar o status de sua solicitação verifique as instruções no seu email.", 5, 1)



			If $status_remot = 1 Then


				_sair()

			Else

				TrayTip("AIT HelpDesk ®", "Pressione F6 para encerra o suporte a qualquer momento.", 5, 1)
				
				TrayTip("AIT HelpDesk ®", "Assitencia Remota Ativada por favor não desconecte ate que seu problema seja resolvido.", 5, 1)
				#Region ### START Koda GUI section ### Form=C:\aithelpdesk\FORMULARIO\Source_Form\aithelpdesk2.kxf
				$AITHELPDESK = GUICreate("Sessão Remota Ativa", 286, 236, 891, 365)
				GUISetBkColor(0xA6CAF0)
				$sair_suporte = GUICtrlCreateButton("Sair", 89, 168, 107, 25)
				$ID_SUPORTE = GUICtrlCreateGroup("ID Suporte", 26, 72, 233, 41)
				$Tickek_id = GUICtrlCreateLabel("TKT ID:", 33, 89, 42, 17)
				$Label1 = GUICtrlCreateLabel("Tipo :", 147, 88, 31, 17)
				$Input1 = GUICtrlCreateInput($random_n, 73, 85, 73, 21)
				$Input2 = GUICtrlCreateInput($inputPriorit, 180, 85, 73, 21)
				GUICtrlCreateGroup("", -99, -99, 1, 1)
				$Pic1 = GUICtrlCreatePic($FILE_1, 36, 8, 164, 60)
				$Pic2 = GUICtrlCreatePic($FILE_6, 200, 8, 44, 28)
				$Button2 = GUICtrlCreateButton("Curtir", 200, 40, 43, 25)
				$Pic3 = GUICtrlCreatePic($FILE_4, 8, 204, 84, 28)
				$Group1 = GUICtrlCreateGroup("ID Local", 27, 120, 233, 41)
				$comp_n = GUICtrlCreateLabel("WRK:", 35, 138, 33, 17)
				$ip_ait = GUICtrlCreateLabel("IP :", 152, 137, 20, 17)
				$Input3 = GUICtrlCreateInput($nomeComputador, 74, 133, 73, 21)
				$Input4 = GUICtrlCreateInput(@IPAddress1 & ", " & @IPAddress2 & ", " & @IPAddress3 & ", " & @IPAddress4, 181, 133, 73, 21)
				GUICtrlCreateGroup("", -99, -99, 1, 1)
				GUISetState(@SW_SHOW)


				While 1


					If ProcessExists("msra.exe") Then; verifica se o processo ex
						;TrayTip("AIT HelpDesk ®", "Assitencia Remota Ativada por favor não desconecte ate que seu problema seja resolvido.", 5, 1)
						;Sleep(2000)
						;TrayTip("AIT HelpDesk ®", "Pressione F4 para encerra o suporte a qualquer momento.", 5, 1)
						;Sleep(2000)


					ElseIf ProcessExists("ait_helpdesk_vnc.exe") Then; verifica se o processo ex
						;TrayTip("AIT HelpDesk ®", "Assitencia Remota Ativada por favor não desconecte ate que seu problema seja resolvido.", 5, 1)
						;Sleep(3000)
						;TrayTip("AIT HelpDesk ®", "Pressione F4 para encerra o suporte a qualquer momento.", 5, 1)
						;Sleep(2000)
						;TrayTip("AIT HelpDesk ®", "Suporte ativo.", 5, 1)
						;Sleep(2000)

					Else
						_sair()
						ExitLoop

					EndIf


					;Sleep(50000)


					$nMsg2 = GUIGetMsg()
					Switch $nMsg2



						Case $GUI_EVENT_CLOSE
							Local $iAnswer = MsgBox(36, "Atenção", "Deseja finalizar o suporte remoto ?")
							;If iAnswer = Sim
							If $iAnswer = 6 Then

								GUIDelete()
								_sair()

								ExitLoop
								Exit
							EndIf

						Case $sair_suporte
							Local $iAnswer = MsgBox(36, "Atenção", "Deseja finalizar o suporte remoto ?")
							;If iAnswer = Sim
							If $iAnswer = 6 Then

								GUIDelete()
								_sair()

								ExitLoop
								Exit
							EndIf

						Case $Button2

							#cs
								; Use no proxy
								HttpSetProxy(1)

								; Use IE defaults for proxy
								HttpSetProxy(0)
								; Use the proxy "www-cache.myisp.net" on port 8080
								HttpSetProxy(2, "www-cache.myisp.net:8080")
							#ce
							ShellExecute("http://www.facebook.com/pages/Bhns/430524047029097", @SW_MAXIMIZE, "", "open") ;http://localhost/osticket/view.php?e=clebe@live.com&t=201504
							TrayTip("AIT HelpDesk ®", "Obrigado " & $inputNome & @LF & " ficamos satisfeitos...." & @LF & " AIT HelpDesk ® agradece.", 5, 1)
							

					EndSwitch


				WEnd




				_sair()

				ExitLoop
				Exit

			EndIf






			;Botão Print Screen
		Case $botaoPrtScreen
			
			
			
			TrayTip("", "", 2, 1)
			$tiket_img_dir = $last_tickets & $random_n & "\img\"
			;MsgBox(4096, "Caminho pasta",$tiket_img_dir)
			;MsgBox(4096, "Caminho pasta",$tiket_img_dir)
			WinSetState("AIT HelpDesk ®", "", @SW_MINIMIZE)
			Sleep(100)
			;$random_a = Random(1111111111111111, 9999999999999999, 1)
			$random_a = $random_a + 1
			$arquivo = "AIT" & @YEAR & @MON & @MDAY & "-IDTK" & $random_a & ".jpg";
			;Tira o PrintScreen

			_ScreenCapture_Capture($tiket_img_dir & $arquivo)
			
			
			_soma_array()
			
			;MsgBox(4096, "Caminho quantidade array",$aPrtScreenSize)
			;MsgBox(4096, "Caminho array",$arrayPrtScreen)
			;Maximiza a janela
			
			WinActivate("AIT HelpDesk ®", "")
			TrayTip("Aviso", "Print screen tirada com suceso." & @LF & "Clique em visualizar para ver a última print screen." & @LF & "Total de : " & $aPrtScreenSize - 1 & " Anexos" & @CRLF & "Total de : " & Int($ticket_full_size / 1048576) & " MB", 2, 1)


			;Botão Visualizar
		Case $botaoVisualizar
			
			;Só carrega a imagem caso o usuário tenha tirado uma printScreen
			If ($imagemSalva = 1) Then
				;Abre uma janela com a imagem salva
				ShellExecute($tiket_img_dir & $arquivo, @SW_MAXIMIZE, "", "open")
			Else
				
				
				Dim $iMsgBoxAnswer
				
				$iMsgBoxAnswer = MsgBox(8244, "Erro ao carregar arquivo", "Nenhuma imagem ou anexo foi carregado, Deseja carregar uma imagem ou arquivo ?")
				Select
					Case $iMsgBoxAnswer = 6 ;Yes

						_anexos()
						
					Case $iMsgBoxAnswer = 7 ;No

				EndSelect
				
				
				
			EndIf



		Case $botaoAnexar
			

			
			_anexos()
			
			
		Case $bt_ferrametas
			
			ShellExecute("http://" & $api_web_front & "view.php?e=" & $NAME_EMAIL & "&t=" & $NAME_LASTTICKET, @SW_MAXIMIZE, "", "open") ;http://localhost/osticket/view.php?e=clebe@live.com&t=201504
			TrayTip("AIT HelpDesk ®", "Obrigado " & $inputNome & @LF & "ficamos satisfeitos...." & @LF & "AIT HelpDesk ® agradece.", 5, 1)

	EndSwitch
	
WEnd

Func _anexos()


	
	; Create a constant variable in Local scope of the message to display in FileOpenDialog.
	Local $sMessage = "Hold down Ctrl or Shift to choose multiple files."

	; Display an open dialog to select a list of file(s).
	Local $sFileOpenDialog = FileOpenDialog($sMessage, @UserProfileDir & "\", "Files (*.docx;*.xlsx;*.doc;*.xls;*.pdf;*.tif;*.txt;*.jpeg;*.gif;*.png;*.jpg;*.bmp)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
	If @error Then
		; Display the error message.
		Dim $iMsgBoxAnswer
		
		$iMsgBoxAnswer = MsgBox(8244, "Erro ao carregar arquivo", "Nenhuma imagem ou anexo foi carregado, Deseja carregar uma imagem ou arquivo ?")
		Select
			Case $iMsgBoxAnswer = 6 ;Yes

				_anexos()
				
			Case $iMsgBoxAnswer = 7 ;No

		EndSelect

		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		;FileChangeDir(@ScriptDir)
	Else
		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		;FileChangeDir(@ScriptDir)
		; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
		;$sFileOpenDialog = StringReplace($sFileOpenDialog, "|", ";")
		; Display the list of selected files.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
		
		
		Local $files_splits = StringSplit($sFileOpenDialog, "|") ; Split the string of days using the delimiter "," and the default flag value.
		#cs
			The array returned will contain the following values:
			$files_splits[1] = "Mon"
			$files_splits[2] = "Tues"
			$files_splits[3] = "Wed"
			...
			$files_splits[7] = "Sun"
		#ce

		For $i = 1 To $files_splits[0] ; Loop through the array returned by StringSplit to display the individual values.
			
			;MsgBox($MB_SYSTEMMODAL, "", "Sequencia : " & $i & " array :" & $files_splits & " Valor :" & $files_splits[0])
			
			
			If $files_splits[0] = 1 Then
				
				$tiket_img_dir = $files_splits[$i]
				$arquivo = '' ; Para limpar memoria anterior
				
				_soma_array()
				
			Else
				
				
				If $i = 1 Then
					
					$tiket_img_dir = $files_splits[$i]

					;MsgBox($MB_SYSTEMMODAL, "","Diretorio : "&$tiket_img_dir&" e "&$files_splits[$i])
					
				ElseIf $i >= 2 Then

					$random_a = $random_a + 1
					;$arquivo = "\AIT" & @YEAR & @MON & @MDAY & "-IDTK" & $random_a & ".jpg";
					$arquivo = "\" & $files_splits[$i]
					
					_soma_array()
					
				Else
					
					
					
				EndIf
				
				
			EndIf
			
			
		Next
		
		WinActivate("AIT HelpDesk ®", "")
		TrayTip("Aviso", "Anexos anexados com suceso." & @LF & "Clique em visualizar para ver o último Anexo." & @LF & "Total de : " & $aPrtScreenSize - 1 & " Anexos" & @CRLF & "Total de : " & Int($ticket_full_size / 1048576) & " MB", 2, 1)


	EndIf

EndFunc   ;==>_anexos


Func _soma_array()

	$imagemSalva = 1
	
	Local $get_file_size = FileGetSize($tiket_img_dir & $arquivo)
	
	;$get_file_size = $get_file_size / 1024
	
	
	If $get_file_size < $get_ticket_anexo_size Then
		
		
		If $ticket_full_size < $get_ticket_ticket_size Then
			
			;MsgBox(64,"Aviso ","Arquivo : "&$tiket_img_dir &  $arquivo&" Total KB: "&$get_file_size  / 1024)
			$arrayPrtScreen[$aPrtScreenSize - 1] = $tiket_img_dir & $arquivo
			$aPrtScreenSize = $aPrtScreenSize + 1
			ReDim $arrayPrtScreen[$aPrtScreenSize]
			$ticket_full_size = $ticket_full_size + $get_file_size ; soma
			
		Else
			
			MsgBox(64, "Aviso ", "O Ticket execede o maximo permetido de : " & Int($get_ticket_ticket_size / 1048576) & " MB de anexos" & @CRLF & "Duvidas abra um ticket com o suporte obrigado!")
			;TrayTip("Aviso", "Anexos anexados com suceso." & @LF & "Clique em visualizar para ver o último Anexo." & @LF & " Total de : " & $aPrtScreenSize - 1, 2, 1)
			
			
		EndIf
		
		
	Else
		
		MsgBox(64, "Aviso ", "O Arquivo : " & $tiket_img_dir & $arquivo & " execede o permetido de : " & Int($get_ticket_anexo_size / 1024) & " KB" & @CRLF & "Tamanho do arquivo : " & Int($get_file_size / 1024) & " KB")
		;TrayTip("Aviso", "Anexos anexados com suceso." & @LF & "Clique em visualizar para ver o último Anexo." & @LF & " Total de : " & $aPrtScreenSize - 1, 2, 1)
		
	EndIf
	
	
EndFunc   ;==>_soma_array


Func urlencode($str) ;U+00D4	 0xD4
	$return = ""

	$str = StringReplace($str, @CRLF, "<br/>")
	$str = StringReplace($str, @LF, "<br/>")
	$str = StringReplace($str, @CR, "<br/>")
	$str = StringReplace($str, Chr(13), "<br/>")
	;$str = StringReplace($str, " ", "&nbsp;")
	$str = StringReplace($str, "Ç", "C")
	$str = StringReplace($str, "ç", "C")
	$str = StringReplace($str, "ã", "a")
	$str = StringReplace($str, "Ã", "A")
	$str = StringReplace($str, "á", "a")
	$str = StringReplace($str, "Ã", "A")
	$str = StringReplace($str, "à", "a")
	$str = StringReplace($str, "À", "A")
	$str = StringReplace($str, "â", "a")
	$str = StringReplace($str, "Â", "A")
	$str = StringReplace($str, "õ", "o")
	$str = StringReplace($str, "é", "e")
	$str = StringReplace($str, "É", "E")
	$str = StringReplace($str, "ó", "o")
	$str = StringReplace($str, "ò", "o")
	$str = StringReplace($str, "Ó", "O")
	$str = StringReplace($str, "Ò", "O")
	$str = StringReplace($str, "ô", "o")
	$str = StringReplace($str, "Ô", "O")
	$return = $str
	$return = _urlencode($return)
	$return = StringReplace($str, "%20", " ")
	
	Return $return
EndFunc   ;==>urlencode


Func _urlencode($string)
	$string = StringSplit($string, "")
	For $i = 1 To $string[0]
		If AscW($string[$i]) < 48 Or AscW($string[$i]) > 122 Then
			$string[$i] = "%" & _StringToHex($string[$i])
		EndIf
	Next
	$string = _ArrayToString($string, "", 1)
	Return $string
EndFunc   ;==>_urlencode









Func _Connect($host, $usr, $pass, $pid)



	Sleep(100)
	MsgBox(0, "Enviado:", $pid)
	;EndIf
	If Not $pid Then _Err("Failed to connect", 0)
	$currentpid = $pid
	$rtn = _Read($pid);Check for Login Success - Prompt
	Sleep(100)
	MsgBox(0, "Recebido:", $rtn)



	;========= Update key ssh =======================
	If StringInStr($rtn, "(y/n)") Then
		_Send($pid, "y" & @CR)
		$rtn = _Read($pid)
	EndIf


	;========= Update key ssh =======================
	If StringInStr($rtn, "yes/no") Then
		_Send($pid, "yes" & @CR)
		$rtn = _Read($pid)
		MsgBox(0, "Recebido:", $rtn)
	EndIf

	;========= Validando login =======================
	If StringInStr($rtn, "Login") Then
		MsgBox(0, "Recebido:", $usr)
		Sleep(100)
		_Send($pid, $usr & @CRLF)
		$rtn = _Read($pid)
		MsgBox(0, "Recebido:", $rtn)
	EndIf


	;========= Validando password ===================
	If StringInStr($rtn, "password") Then
		MsgBox(0, "Recebido:", $pass)
		Sleep(100)
		_Send($pid, $pass & @CRLF)
		$rtn = _Read($pid)
		MsgBox(0, "Recebido:", $rtn)
	EndIf


	_Send($pid, "" & @CR)
	$rtn = _Read($pid)
	Sleep(100)
	MsgBox(0, "Enviado Enter:", $rtn)

	_Send($pid, "ping 8.8.8.8" & @CR)
	$rtn = _Read($pid)
	Sleep(100)
	MsgBox(0, "ping 8.8.8.8:", $rtn)


	If StringInStr($rtn, "Access denied") Or StringInStr($rtn, "FATAL") Then _Err($rtn, $pid)
	Return $pid
EndFunc   ;==>_Connect

Func _Read($pid)
	If Not $pid Then Return -1
	Local $dataA
	Local $dataB
	Do
		$dataB = $dataA
		Sleep(100)
		$dataA &= StdoutRead($pid)
		If @error Then ExitLoop
	Until $dataB = $dataA And $dataA And $dataB
	FileWriteLine(@ScriptDir & "\log.txt", $dataA & @CRLF)
	Return $dataA
EndFunc   ;==>_Read

Func _Send($pid, $cmd)
	StdinWrite($pid, $cmd)
EndFunc   ;==>_Send

Func _Err($data, $pid)
	If $data And $data <> -1 Then MsgBox(0, "An Error has Occured", $data, 5)
	_Exit($pid)
EndFunc   ;==>_Err

Func _Exit($pid)
	ProcessClose($pid)
EndFunc   ;==>_Exit




Func enabled_vnc_assitencie()


	If ProcessExists("ait_helpdesk_vnc.exe") Then; verifica se o processo ex

		TrayTip("AIT HelpDesk ®", "Fechando sessões anteriores abertas e gerando uma nova ID.", 5, 1)
		

		;If not ProcessExists($gf_g_1) Then; verifica se o processo, existe não faz nada, se não existe cria um novo
		;SplashTextOn("Gerando uma nova Chave","",300, 50,-1, -1, 1, "Arial", 18)
		;controlSetText("Gerando uma nova Chave", "", "Static1", " Chave nova sessão AITKEY"&$random_n) ;o 5 foi usado para fazer uma contagem regressiva
		;Sleep(3000)
		;Splashoff()
		ProcessClose("ait_helpdesk_vnc.exe")
	ElseIf ProcessExists("msra.exe") Then; verifica se o processo ex

		TrayTip("AIT HelpDesk ®", "Fechando sessões anteriores abertas e gerando uma nova ID.", 5, 1)
		
		;If not ProcessExists($gf_g_1) Then; verifica se o processo, existe não faz nada, se não existe cria um novo
		;SplashTextOn("Gerando uma nova Chave","",300, 50,-1, -1, 1, "Arial", 18)
		;controlSetText("Gerando uma nova Chave", "", "Static1", " Chave nova sessão AITKEY"&$random_n) ;o 5 foi usado para fazer uma contagem regressiva
		;Sleep(3000)
		;Splashoff()
		ProcessClose("msra.exe")

	EndIf

	TrayTip("AIT HelpDesk ®", " Configurando suporte remoto....," & @LF & " Aguarde.", 5, 1)
	RegWrite('HKEY_CURRENT_USER \ Software\TightVNC\Server', 'ExtraPorts', 'REG_SZ', '')
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'AcceptRfbConnections', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'AcceptHttpConnections', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'AllowLoopback', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'AlwaysShared', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'BlockLocalInput', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'BlockRemoteInput', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'QueryTimeout', 'REG_DWORD', 30)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'QueryAcceptOnTimeout', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'LocalInputPriorityTimeout', 'REG_DWORD', 3)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'LocalInputPriority', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'BlockRemoteInput', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'IpAccessControl', 'REG_SZ', '')
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'RfbPort', 'REG_DWORD', 6064)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'HttpPort', 'REG_DWORD', 5800)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'DisconnectAction', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'UseVncAuthentication', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'UseControlAuthentication', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'LoopbackOnly', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'LogLevel', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'EnableFileTransfers', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'RemoveWallpaper', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'UseMirrorDriver', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'EnableUrlParams', 'REG_DWORD', 1)
	;RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'Password', 'REG_BINARY', 'ê2ÚMr½‡Þ') ; key : 88296064
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'Password', 'REG_BINARY', 'Ojt~ƒ5ç') ; 70x7
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'PasswordViewOnly', 'REG_BINARY', 'Ô-€_	K#')
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'NeverShared', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'DisconnectClients', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'DisconnectClients', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'VideoRecognitionInterval', 'REG_DWORD', 3000)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'GrabTransparentWindows', 'REG_DWORD', 1)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'SaveLogToAllUsersPath', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'RunControlInterface', 'REG_DWORD', 0)
	RegWrite('HKEY_CURRENT_USER\Software\TightVNC\Server', 'VideoClasses', 'REG_SZ', '')


	#CS
		Registro Windows liberando VNC porta 6064

		Senha padrão 88296064
		Senha viewr 25190422

		Windows Registry Editor Version 5.00

		[HKEY_CURRENT_USER\Software\TightVNC]
		[HKEY_CURRENT_USER\Software\TightVNC\Server]
		"ExtraPorts"=""
		"QueryTimeout"=dword:0000001e
		"QueryAcceptOnTimeout"=dword:00000000
		"LocalInputPriorityTimeout"=dword:00000003
		"LocalInputPriority"=dword:00000000
		"BlockRemoteInput"=dword:00000000
		"BlockLocalInput"=dword:00000000
		"IpAccessControl"=""
		"RfbPort"=dword:000017b0
		"HttpPort"=dword:000016a8
		"DisconnectAction"=dword:00000000
		"AcceptRfbConnections"=dword:00000001
		"UseVncAuthentication"=dword:00000001
		"UseControlAuthentication"=dword:00000000
		"LoopbackOnly"=dword:00000000
		"AcceptHttpConnections"=dword:00000000
		"LogLevel"=dword:00000000
		"EnableFileTransfers"=dword:00000001
		"RemoveWallpaper"=dword:00000000
		"UseMirrorDriver"=dword:00000001
		"EnableUrlParams"=dword:00000001
		"Password"=hex:ea,32,da,4d,72,bd,87,de
		"PasswordViewOnly"=hex:d4,02,2d,80,5f,09,4b,23
		"AlwaysShared"=dword:00000000
		"NeverShared"=dword:00000000
		"DisconnectClients"=dword:00000001
		"DisconnectClients"=dword:000003e8
		"AllowLoopback"=dword:00000000
		"VideoRecognitionInterval"=dword:00000bb8
		"GrabTransparentWindows"=dword:00000001
		"SaveLogToAllUsersPath"=dword:00000000
		"RunControlInterface"=dword:00000000
		"VideoClasses"=""

	#CE

	;Run(@ComSpec & $remote_vnc_file & " -run ", @TempDir, @SW_HIDE)
	;ShellExecute($remote_vnc_file & " -run ", , "", "")

	open_ports_vnc()

	Run($remote_vnc_file & " -run ", @SystemDir, "", 0x1 + 0x8)
	; Run a command prompt as the other user.
	;RunAs($sUserName, @ComputerName, $sPassword, 0, @ComSpec, @SystemDir)
	;Run(@ComSpec & $remote_vnc_file & " -run ", @SystemDir, "", 0x1 + 0x8)


	TrayTip("AIT HelpDesk ®", "Remoto VNC habilitado ID-VNC_" & $random_n, 5, 1)

	

EndFunc   ;==>enabled_vnc_assitencie


Func open_ports_vnc()

	TrayTip("AIT HelpDesk ®", " Liberando Portas firewall....," & @LF & " Aguarde...", 5, 1)
	; Run a command prompt as the other user.

	Local $login_local = StringReplace(@LogonServer, "\\", "")
	
	MsgBox(64, "Combinar", $login_local & @CRLF & @ComputerName)
	
	If $adm_force = 0 Then
		
		RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @TempDir, @SW_HIDE)
		RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @TempDir, @SW_HIDE)

		;Run(' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @TempDir, @SW_HIDE)
		;Run(' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @TempDir, @SW_HIDE)
		;Run(@ComSpec & "/k " & 'netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program="' &  $remote_vnc_file & '"' & ' enable=yes')
		;Run(@ComSpec & "/k " & 'netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes')

		;MsgBox(64, "Forcer", " Modo forçado ")
		
	ElseIf $login_local = @ComputerName Then
		
		;RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @TempDir, @SW_HIDE)
		;RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @TempDir, @SW_HIDE)


		Run(@ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @TempDir, @SW_HIDE)
		Run(@ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @TempDir, @SW_HIDE)
		;RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @SystemDir)
		;RunAs($sUserName, $sDomains_Work, $sPassword, 0, @ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @SystemDir)

		;	MsgBox(64, "Interno", " Login ")
		
	Else
		
		Run(' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program=' & '"' & $remote_vnc_file & '"' & ' enable=yes', @TempDir, @SW_HIDE)
		Run(' /c netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes', @TempDir, @SW_HIDE)
		; Run(@ComSpec & "/k " & 'netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC"' & ' dir=in action=allow program="' &  $remote_vnc_file & '"' & ' enable=yes')
		; Run(@ComSpec & "/k " & 'netsh advfirewall firewall add rule name=' & '"Remoto AIT-Helpdesk VNC Port"' & ' dir=in action=allow protocol=TCP localport=6064' & ' enable=yes')


	EndIf

	
	
EndFunc   ;==>open_ports_vnc



;======================================================================================================


Func open_ports_rdp()

	TrayTip("AIT HelpDesk ®", " Liberando Portas firewall....," & @LF & " Aguarde.", 5, 1)
	; Grupo remote Windows PT  Assistência Remota
	Run(@ComSpec & ' /c netsh advfirewall firewall set rule group=' & '"Área de Trabalho Remota"' & ' new enable=Yes', @TempDir, @SW_HIDE)
	Run(@ComSpec & ' /c netsh advfirewall firewall set rule group=' & '"Assistência Remota"' & ' new enable=Yes', @TempDir, @SW_HIDE)
	;Run(@ComSpec & ' /c netsh advfirewall firewall add rule name=' & '"AIT-HelpDesk Open Port 22 Plink"' & ' dir=in action=allow program=' & '"' & $Plink_File & '"' & ' enable=yes', @TempDir, @SW_HIDE)
	;Grupo remote Windows ENG  Assistência Remota
	Run(@ComSpec & ' /c netsh advfirewall firewall set rule group=' & '"remote desktop"' & ' new enable=Yes', @TempDir, @SW_HIDE)
	;Run(@ComSpec & ' /c netsh advfirewall firewall set rule group='&'"Assistência Remota"'&' new enable=Yes', @TempDir, @SW_HIDE)

EndFunc   ;==>open_ports_rdp



Func enabled_rdp_assitencie()


	If ProcessExists("msra.exe") Then; verifica se o processo ex

		TrayTip("AIT HelpDesk ®", "Fechando sessões anteriores abertas e gerando uma nova ID.", 5, 1)


		;If not ProcessExists($gf_g_1) Then; verifica se o processo, existe não faz nada, se não existe cria um novo
		;SplashTextOn("Gerando uma nova Chave","",300, 50,-1, -1, 1, "Arial", 18)
		;controlSetText("Gerando uma nova Chave", "", "Static1", " Chave nova sessão AITKEY"&$random_n) ;o 5 foi usado para fazer uma contagem regressiva
		;Sleep(3000)
		;Splashoff()
		ProcessClose("msra.exe")
	ElseIf ProcessExists("ait_helpdesk_vnc.exe") Then; verifica se o processo ex

		TrayTip("AIT HelpDesk ®", "Fechando sessões anteriores abertas e gerando uma nova ID.", 5, 1)
		

		;If not ProcessExists($gf_g_1) Then; verifica se o processo, existe não faz nada, se não existe cria um novo
		;SplashTextOn("Gerando uma nova Chave","",300, 50,-1, -1, 1, "Arial", 18)
		;controlSetText("Gerando uma nova Chave", "", "Static1", " Chave nova sessão AITKEY"&$random_n) ;o 5 foi usado para fazer uma contagem regressiva
		;Sleep(3000)
		;Splashoff()
		ProcessClose("ait_helpdesk_vnc.exe")
	EndIf
	$random_rdp = Random(111111, 999999, 1)
	$random_KEY = Random(111111, 999999, 1)
	$file_rdp = "AITRDP_" & @YEAR & @MON &"_"&$random_n & ".msrcincident"
	FileDelete($folder_convicts & "*.*") ; Deleta todos os convites antigos na maquina
	
	Run(@ComSpec & " /c msra.exe /saveasfile " & $folder_convicts & $file_rdp & " " & "AITKEY" & $random_KEY, @TempDir, @SW_HIDE)
	;Run(@ComSpec & ' /c msra.exe /offerra ' & $ComputerName, @TempDir, @SW_HIDE)
	;Run("msra.exe /saveasfile "&$folder_aithelpdesk&" "&$random_n)
	;Run(@SystemDir &"\msra.exe /saveasfile path "&$random_n, @SystemDir, "", "")
	;ShellExecute(@SystemDir &"\msra.exe  /saveasfile  path  "&$random_n)
	;msra.exe /saveasfile path cleber

	;WinWait("Windows Remote Assistance")
	;WinActivate("Windows Remote Assistance")
	;ControlClick("Windows Remote Assistance", "", "[CLASS:Button; TEXT:Yes; INSTANCE:1]")
	TrayTip("AIT HelpDesk ®", "Windows Remote Assistance ativado chave AITKEY" & $random_n, 5, 1)
	
    
	;$json_attachments_rdp = '{"'"AITKEY"& $random_KEY"'":"data:image/jpeg;base64,' & $file_encode64 & '"}'


EndFunc   ;==>enabled_rdp_assitencie


Func _About()

EndFunc   ;==>_About


Func _sair()


	If ProcessExists("msra.exe") Then; verifica se o processo ex
		TrayTip("AIT HelpDesk ®", "Fechando Assitencia Remota .", 5, 1)
		
		ProcessClose("msra.exe")
		disconect_VPN()
		report_close()

	ElseIf ProcessExists("ait_helpdesk_vnc.exe") Then; verifica se o processo ex
		TrayTip("AIT HelpDesk ®", "Fechando Assitencia Remota VNC.", 5, 1)
		
		ProcessClose("ait_helpdesk_vnc.exe")
		disconect_VPN()
		report_close()
	Else
		TrayTip("AIT HelpDesk ®", "Fim de sessão de suporte.", 5, 1)



		TrayTip("AIT HelpDesk ®", " Agradeçemos sua preferência," & @LF & " obrigado equipe AIT Helpdesk...", 5, 1)
		

		Exit


	EndIf
	GUIDelete()

	
EndFunc   ;==>_sair




Func report_close()




EndFunc   ;==>report_close
#cs
	netsh advfirewall firewall set rule group="Área de Trabalho Remota" new enable=Yes

	netsh advfirewall firewall set rule group="remote desktop" new enable=Yes

	netsh advfirewall firewall add rule name="AIT-HelpDesk Open Port 22 Plink" dir=in action=allow program="C:\MyApp\MyApp.exe" enable=yes


	netsh advfirewall firewall set rule group="Área de Trabalho Remota" new enable=Yes profile=all
	netsh advfirewall firewall set rule group="remote desktop" new enable=Yes profile=private
	netsh advfirewall firewall set rule group="remote desktop" new enable=Yes



	netsh advfirewall firewall add rule name ="AIT-HelpDesk Open Port Plink" dir = em ação = permitir protocolo = TCP localport = 22
	netsh advfirewall firewall add rule name ="abrir porta 80" dir = em ação = permitir protocolo = TCP localport = 80

	netsh advfirewall firewall add rule name="AIT-HelpDesk Open Port 22 Plink" dir=in action=allow protocol=TCP localport=22


#ce



;Code below was generated by: 'File to Base64 String' Code Generator v1.11 Build 2013-02-19

Func rasdial_conect()

	RunWait("rasdial ""VPN Connection"" username password", @WorkingDir, @SW_HIDE)
EndFunc   ;==>rasdial_conect


Func rasdial_disconect()

	RunWait("rasdial ""VPN Connection"" /disconnect", @WorkingDir, @SW_HIDE)
EndFunc   ;==>rasdial_disconect

Func criar_profile_vpn()

	;==================caminho profile==========================
	TrayTip("AIT HelpDesk ®", " Criando profile VPN....," & @LF & " Aguarde...", 5, 1)
	If DirGetSize($folder_pbk) = -1 Then
		DirCreate($folder_pbk)
	ElseIf DirGetSize($folder_bkp_pbk) = -1 Then
		DirCreate($folder_bkp_pbk)
		DirCopy($folder_pbk, $folder_bkp_pbk, 1)
		DirRemove($folder_pbk, 1)
		DirCreate($folder_pbk)
	Else
	EndIf

	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "LAST", "1")

	;=====================iniWritre erro adaptado fileWitre=============
	FileWriteLine($folder_pbk & "\rasphone.pbk", "[AITHELPDESK]") ; Name Profile
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Encoding=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PBVersion=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Type=2")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "AutoLogon=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "UseRasCredentials=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "LowDateTime=-210050432")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "HighDateTime=30285916")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DialParamsUID=387729268")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Guid=4A93C8C09DEF224893C736A9070F1ED2")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "VpnStrategy=2")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ExcludedProtocols=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "LcpExtensions=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DataEncryption=256")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "SwCompression=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "NegotiateMultilinkAlways=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "SkipDoubleDialDialog=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DialMode=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "OverridePref=15")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "RedialAttempts=3")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "RedialSeconds=60")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IdleDisconnectSeconds=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "RedialOnLinkFailure=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CallbackMode=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CustomDialDll=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CustomDialFunc=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CustomRasDialDll=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ForceSecureCompartment=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DisableIKENameEkuCheck=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "AuthenticateServer=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ShareMsFilePrint=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "BindMsNetClient=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "SharedPhoneNumbers=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "GlobalDeviceSettings=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PrerequisiteEntry=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PrerequisitePbk=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredPort=VPN1-0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredDevice=WAN Miniport (IKEv2)")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredBps=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredHwFlow=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredProtocol=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredCompression=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredSpeaker=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreferredMdmProtocol=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreviewUserPw=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreviewDomain=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreviewPhoneNumber=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ShowDialingProgress=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ShowMonitorIconInTaskBar=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CustomAuthKey=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "AuthRestrictions=544")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpPrioritizeRemote=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpInterfaceMetric=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpHeaderCompression=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpAddress=0.0.0.0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpDnsAddress=0.0.0.0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpDns2Address=0.0.0.0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpWinsAddress=0.0.0.0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpWins2Address=0.0.0.0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpAssign=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpNameAssign=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpDnsFlags=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpNBTFlags=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "TcpWindowSize=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "UseFlags=2")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpSecFlags=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "IpDnsSuffix=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6Assign=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6Address=::")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6PrefixLength=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6PrioritizeRemote=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6InterfaceMetric=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6NameAssign=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6DnsAddress=::")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6Dns2Address=::")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6Prefix=0000000000000000")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Ipv6InterfaceId=5C409CF71F1DE54D")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DisableClassBasedDefaultRoute=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DisableMobility=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "NetworkOutageTime=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ProvisionType=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PreSharedKey=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", @CRLF)
	FileWriteLine($folder_pbk & "\rasphone.pbk", "NETCOMPONENTS=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ms_msclient=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "ms_server=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", @CRLF)
	FileWriteLine($folder_pbk & "\rasphone.pbk", "MEDIA=rastapi")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Port=VPN1-0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Device=WAN Miniport (IKEv2)")
	FileWriteLine($folder_pbk & "\rasphone.pbk", @CRLF)
	FileWriteLine($folder_pbk & "\rasphone.pbk", "DEVICE=vpn")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PhoneNumber=" & $host_vpn) ; Host
	FileWriteLine($folder_pbk & "\rasphone.pbk", "AreaCode=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CountryCode=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "CountryID=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "UseDialingRules=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "Comment=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "FriendlyName=")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "LastSelectedPhone=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "PromoteAlternates=0")
	FileWriteLine($folder_pbk & "\rasphone.pbk", "TryNextAlternateOnFail=1")
	FileWriteLine($folder_pbk & "\rasphone.pbk", @CRLF)
	FileWriteLine($folder_pbk & "\rasphone.pbk", @CRLF)
	conect_VPN()

	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "BACKUPFILE", "1")

EndFunc   ;==>criar_profile_vpn

Func disconect_VPN()
	Run("rasdial " & $rasName & " /disconnect", "", @SW_DISABLE)

	Sleep(500)


	If DirGetSize($folder_bkp_pbk) > 1 Then
		;MsgBox(64, "existe pasta backup", $folder_bkp_pbk)
		DirRemove($folder_pbk, 1)
		DirCopy($folder_bkp_pbk, $folder_pbk, 1)
		DirRemove($folder_bkp_pbk, 1)
	Else
		;MsgBox(64, "nao existe pasta backup", $folder_bkp_pbk)
		DirRemove($folder_pbk, 1)

	EndIf

	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "LAST", "0")
	IniWrite($folder_config & "aithelpdesk.ini", "SESSION", "BACKUPFILE", "0")

EndFunc   ;==>disconect_VPN

Func conect_VPN()
	;$pid = Run($Plink_File & " -ssh " & $host & " -P " & $port, @SystemDir, "", 0x1 + 0x8)
	;$pid = Run("rasdial "&$rasName&" /Disconnect","",@SW_DISABLE)
	;$pid = Run("rasdial", @SystemDir, @SW_DISABLE, $STDERR_CHILD + $STDOUT_CHILD)
	Run("rasdial " & $rasName & " " & $email_vpn & " " & $passwd_vpn, "", @SW_DISABLE)
	
EndFunc   ;==>conect_VPN




Func _STRING_1($bSaveBinary = False)
	Local $Base64String
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'CQB9Fo/qo+3LXVHHHs8zndYymdezr8Wx+NhZwwBl9QqhltVZYfT2/wChdkP+nZ/gV22LkGyQZ0015VF3RMaw3GymtxyWtbeS3+ca3+bbb+9sj2K7h4gx2hjAA1oAaBwAOAnQiQT4rZEECuj/AP/S6x7GvEOEhC+xY37gRklfaWrCumuv6DYU0kklMX1seIeJCF9ixf3AjpJaJ1YNpqYIa0AJ/Sr/AHQpJJIUABxokkkkpia6zy0Jw1reBCdJJSkkkklKIBSSSSUpJJJJSkkkklP/2f/tEvhQaG90b3Nob3AgMy4wADhCSU0EJQAAAAAAEAAAAAAAAAAAAAAAAAAAAAA4QklNBC8AAAAAAEr//wEASAAAAEgAAAAAAAAAAAAAANACAABAAgAAAAAAAAAAAAAYAwAAZAIAAAABwAMAALAEAAABAA8nAQBsbHVuAAAAAAAAAAAAADhCSU0D7QAAAAAAEABIAAAAAQACAEgAAAABAAI4QklNBCYAAAAAAA4AAAAAAAAAAAAAP4AAADhCSU0EDQAAAAAABAAAAJA4QklNBBkAAAAAAAQAAAAeOEJJTQPzAAAAAAAJAAAAAAAAAAABADhCSU0ECgAAAAAAAQAAOEJJTScQAAAAAAAKAAEAAAAAAAAAAjhCSU0D9QAAAAAASAAvZmYAAQBsZmYABgAAAAAAAQAvZmYAAQChmZoABgAAAAAAAQAyAAAAAQBaAAAABgAAAAAAAQA1AAAAAQAtAAAABgAAAAAAAThCSU0D+AAAAAAAcAAA/////////////////////////////wPoAAAAAP////////////////////////////8D6AAAAAD/////////////////////////////A+gAAAAA/////////////////////////////wPoAAA4QklNBAAAAAAAAAIABjhCSU0EAgAAAAAADgAAAAAAAAAAAAAAAAAAOEJJTQQwAAAAAAAHAQEBAQEBAQA4QklNBC0AAAAAABIABAAAAAkAAAAKAAAACAAAAAc4QklNBAgAAAAAABAAAAABAAACQAAAAkAAAAAAOEJJTQQeAAAAAAAEAAAAADhCSU0EGgAAAAADPQAAAAYAAAAAAAAAAAAAADwAAACkAAAABABsAG8AZwBvAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAACkAAAAPAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAABAAAAABAAAAAAAAbnVsbAAAAAIAAAAGYm91bmRzT2JqYwAAAAEAAAAAAABSY3QxAAAABAAAAABUb3AgbG9uZwAAAAAAAAAATGVmdGxvbmcAAAAAAAAAAEJ0b21sb25nAAAAPAAAAABSZ2h0bG9uZwAAAKQAAAAGc2xpY2VzVmxMcwAAAAFPYmpjAAAAAQAAAAAABXNsaWNlAAAAEgAAAAdzbGljZUlEbG9uZwAAAAAAAAAHZ3JvdXBJRGxvbmcAAAAAAAAABm9yaWdpbmVudW0AAAAMRVNsaWNlT3JpZ2luAAAADWF1dG9HZW5lcmF0ZWQAAAAAVHlwZWVudW0AAAAKRVNsaWNlVHlwZQAAAABJbWcgAAAABmJvdW5kc09iamMAAAABAAAAAAAAUmN0MQAAAAQAAAAAVG9wIGxvbmcAAAAAAAAAAExlZnRsb25nAAAAAAAAAABCdG9tbG9uZwAAADwAAAAAUmdodGxvbmcAAACkAAAAA3VybFRFWFQAAAABAAAAAAAAbnVsbFRFWFQAAAABAAAAAAAATXNnZVRFWFQAAAABAAAAAAAGYWx0VGFnVEVYVAAAAAEAAAAAAA5jZWxsVGV4dElzSFRNTGJvb2wBAAAACGNlbGxUZXh0VEVYVAAAAAEAAAAAAAlob3J6QWxpZ25lbnVtAAAAD0VTbGljZUhvcnpBbGlnbgAAAAdkZWZhdWx0AAAACXZlcnRBbGlnbmVu'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'j2K7h4gx2hjAA1oAaBwAOAnQiQT4rZEECuj/AP/S6x7GvEOEhC+xY37gRklfaWrCumuv6DYU0kklMX1seIeJCF9ixf3AjpJaJ1YNpqYIa0AJ/Sr/AHQpJJIUABxokkkkpia6zy0Jw1reBCdJJSkkkklKIBSSSSUpJJJJSkkkklP/2QA4QklNBCEAAAAAAFUAAAABAQAAAA8AQQBkAG8AYgBlACAAUABoAG8AdABvAHMAaABvAHAAAAATAEEAZABvAGIAZQAgAFAAaABvAHQAbwBzAGgAbwBwACAAQwBTADMAAAABADhCSU0PoAAAAAABDG1hbmlJUkZSAAABADhCSU1BbkRzAAAA4AAAABAAAAABAAAAAAAAbnVsbAAAAAMAAAAAQUZTdGxvbmcAAAAAAAAAAEZySW5WbExzAAAAAU9iamMAAAABAAAAAAAAbnVsbAAAAAIAAAAARnJJRGxvbmcrW+YwAAAAAEZyR0Fkb3ViQGIAAAAAAAAAAAAARlN0c1ZsTHMAAAABT2JqYwAAAAEAAAAAAABudWxsAAAABAAAAABGc0lEbG9uZwAAAAAAAAAAQUZybWxvbmcAAAAAAAAAAEZzRnJWbExzAAAAAWxvbmcrW+YwAAAAAExDbnRsb25nAAAAAAAAOEJJTVJvbGwAAAAIAAAAAAAAAAA4QklND6EAAAAAABxtZnJpAAAAAgAAABAAAAABAAAAAAAAAAEAAAAAOEJJTQQGAAAAAAAHAAgAAAABAQD/4Q/NaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/PiA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA0LjEtYzAzNiA0Ni4yNzY3MjAsIE1vbiBGZWIgMTkgMjAwNyAyMjo0MDowOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIiB4bWxuczp4YXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhhcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1sbnM6dGlmZj0iaHR0cDovL25zLmFkb2JlLmNvbS90aWZmLzEuMC8iIHhtbG5zOmV4aWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vZXhpZi8xLjAvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIGRjOmZvcm1hdD0iaW1hZ2UvanBlZyIgeGFwOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1MzIFdpbmRvd3MiIHhhcDpDcmVhdGVEYXRlPSIyMDEzLTAzLTA5VDE3OjA3OjIxLTAzOjAwIiB4YXA6TW9kaWZ5RGF0ZT0iMjAxMy0wMy0wOVQxNzowNzoyMS0wMzowMCIgeGFwOk1ldGFkYXRhRGF0ZT0iMjAxMy0wMy0wOVQxNzowNzoyMS0wMzowMCIgeGFwTU06RG9jdW1lbnRJRD0idXVpZDpCNzMzQkNGM0Y0ODhFMjExOTVCQUY1QTBGMEUxNDMxNCIgeGFwTU06SW5zdGFuY2VJRD0idXVpZDpCODMzQkNGM0Y0ODhFMjExOTVCQUY1'
	$Base64String &= 'QTBGMEUxNDMxNCIgdGlmZjpPcmllbnRhdGlvbj0iMSIgdGlmZjpYUmVzb2x1dGlvbj0iNzIwMDAwLzEwMDAwIiB0aWZmOllSZXNvbHV0aW9uPSI3MjAwMDAvMTAwMDAiIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiIHRpZmY6TmF0aXZlRGlnZXN0PSIyNTYsMjU3LDI1OCwyNTksMjYyLDI3NCwyNzcsMjg0LDUzMCw1MzEsMjgyLDI4MywyOTYsMzAxLDMxOCwzMTksNTI5LDUzMiwzMDYsMjcwLDI3MSwyNzIsMzA1LDMxNSwzMzQzMjtCRUVDMkNBMjVDRUQxRjE4MDY1MDFBNkE0Q0VGODUzMiIgZXhpZjpQaXhlbFhEaW1lbnNpb249IjE2NCIgZXhpZjpQaXhlbFlEaW1lbnNpb249IjYwIiBleGlmOkNvbG9yU3BhY2U9IjEiIGV4aWY6TmF0aXZlRGlnZXN0PSIzNjg2NCw0MDk2MCw0MDk2MSwzNzEyMSwzNzEyMiw0MDk2Miw0MDk2MywzNzUxMCw0MDk2NCwzNjg2NywzNjg2OCwzMzQzNCwzMzQzNywzNDg1MCwzNDg1MiwzNDg1NSwzNDg1NiwzNzM3NywzNzM3OCwzNzM3OSwzNzM4MCwzNzM4MSwzNzM4MiwzNzM4MywzNzM4NCwzNzM4NSwzNzM4NiwzNzM5Niw0MTQ4Myw0MTQ4NCw0MTQ4Niw0MTQ4Nyw0MTQ4OCw0MTQ5Miw0MTQ5Myw0MTQ5NSw0MTcyOCw0MTcyOSw0MTczMCw0MTk4NSw0MTk4Niw0MTk4Nyw0MTk4OCw0MTk4OSw0MTk5MCw0MTk5MSw0MTk5Miw0MTk5Myw0MTk5NCw0MTk5NSw0MTk5Niw0MjAxNiwwLDIsNCw1LDYsNyw4LDksMTAsMTEsMTIsMTMsMTQsMTUsMTYsMTcsMTgsMjAsMjIsMjMsMjQsMjUsMjYsMjcsMjgsMzA7OTFDMUUzMzVBN0VGMEQyNzE1OUUyRTlCRDI2M0Q1MTciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIgcGhvdG9zaG9wOkhpc3Rvcnk9IiI+IDx4YXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ1dWlkOjJGOTJCMzEwRjQ4OEUyMTE5NUJBRjVBMEYwRTE0MzE0IiBzdFJlZjpkb2N1bWVudElEPSJ1dWlkOjkxNkE2NERERjE4OEUyMTE5NUJBRjVBMEYwRTE0MzE0Ii8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
	$Base64String &= 'ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'uq6uMgd1zctNcNgDVIaKBgVffODdTw4HZuxKWUifPZWfNV8aW/4BYqNaWjSXnVoqKyuZlAFi1PckWF5g+73svj7nve/SJ2wxLCh/pSHU5HzCoAfk/wA+uYH98J7mnbeSPbD2ks7kibc76XcblRT+xs08G3V810yTXEjqAKFrapI0ipnOkMVlsZ1BsJMpHKF/hUlPRTyLpEtNSVDrFEp4B+0gkjj/AOChfYc99eWX2DnV71I6Wm5QrOvpr+CUfbrXWf8AmoPsE+/3Xfu8vuX92La+XLy6177ylfS7XJX4jb4ubJ8Y0LBN9MhwT9K1QfibF3Rt6fP7Ayr0KF8tt54N04fSLyfe4MtUSxRcE+arxrVECWIOqUc2v7DftVzH/VnnjZr15NNrK/gS+miUhQT8lfQ5+S9S19/D2eb3p+7D7jcvWdsJeYNugG6WQoC31FgGlZEBB/UntTc2yUp3TCppUHZK/ksfI2LffWGN27V1olnpaaCEI8moiyKF4Jv+PfXbknchcWqRM2QOvjA5usDDcs4XBPVbn/Cr3ncfwbP9cJ8hv/c7pf2Gvdr+02L/AEs3+GPoQe2f9nvP2xf4JOtS/aDMu6MAykqwytGQykgg+ZeQRyD7iKH+1j+0dSbL/Zv9nW5h8QPhB0J82/jttHD957Bffs+xFzlXsondu+dsnDVm4aXFRZSUf3O3Lt7+IfdDC03pqvOsfi9AXU+qadn2Lb9826JL+38Qx109zLSoFfhYV4DjXqJ913i92fcJWs59AemrtU1pWnxA04nh0QP5f/yVaHp7rjuftbae3pcPt7rjYO8t6xQfxnceQWGn21ha7LcyZXJVskoSOk/ts1/z7D+8ckrZ217dwxUjijZuLH4QT5k9He1c3G6uLS2leskjqvADiaeQ61qPcZdSD0ff4nfEsfI7IY7FUdG0tTVINTLU10ZZi5W9oZ1A+n4AHsQbRtH7yZUVcn7eiTc90+gVmJwOr0Ojv5KG7ti7vw+5KSlmp11U0sxNVkZg6JIJVFp5ZQvP9APY7sOSJreZJQCP29A285vimiaMn/B1tw/HzYtV1919iNvVd/LR0kEJuSf83Gq835vx7l/brc29ukZ4gdRhfTCedpBwPQ6+1/SLr3v3Xuve/de697917qkz+YTgczkO1PkhhanB5XcGd7c/lt7z66+NmMxuKrMtkNx7/wAPvbdWf7d2Dsygo6Opq8pvnP4Ku2vXfYUhasraDGsyRSJSSNHHfNMUjXu7RmNmln2h47cAElnDs0qIACS7AxmgyQvA6TQf8syRrZ7VIJFWODdVeckgBUKKInYk0CAiQVOAW49wqAS9YVfRPTn8xfrvtnrHE7Rquz+o+ltpdE4bpbpnNdRdHdn9h5rY+bxXWJ6527Pm9xxD5NR9uZumi3HBDkZawDFUNaiGGCVoiz6Ntt2/my0vrNYzNBEsIiiMUMjlCI/DWrf4x4pHiAMT2qwwDQx+rG433K91ZXbOIZ5WmMsokljQODJrai/oeGDoJUDuZeJFb5P4Zv7/AEOfwf8AiSf6Uf8ARn/DP4v9zH4/7/f3W+1/iX3fg8Oj+8P7vk8em3q0249yXouvoPD1/wCO+DSv9PTStf8ATdR1rtvrtej/ABTxa0/oauFP9L1//9I9X817uo9V/FbM7XxtWKfcfcmUg6/o0Q/5Qm3HifJ7yq1U2VqWXEUox0v1I/iS2H5Aw+9dzqeVfau82u2m07lvMotFpx8EgvcN/pTGvgt/zWH2jJ3+6M9j192vvZ7LzVudmZOWuSrR93kJ+A3gYQbbGTxDrcSG8j4A/RNU+Tam3vk319efSYzOydmbjrYMluHaO2M9kaaFKamr8zgMVlK2np45ZJkp4KqupJ54oUmmdwisFDOTa5Ps427mHf8AaIWt9p3y8tYGbUVhmkjUsQBqIRlBagArStAB5DqOOcPZ32j9w9xg3jn/ANrOXN83aKAQpPuG2WV7MkKs7rEslzDI6xB5JHEYYKGd2Aqx'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'MdZqT+VV8Z9vbF3L1jt7qHHUGzd61WCr910UW5d8yZTNVO2siuVwYqtzz7mm3QtLjMgplip0rUgBZrpZmB2nJezRWs1lDYAW0pUsNT1Ok1WratVAc0rT5dafnDeZbmG8lvybiMMFOlKDUKN26dORitK9D7s74ebX2f2buztbB4VsXvDfv8D/AL519NlcyKPPPgKb7TFVlXgXyDbfTJUlKTH93HSx1LoSHkYE+zSDY4Le8nvYotNxLp1kE0bTgErXTUDFaV9T0WT7zPPaQWUsmqCKugUFV1ZIDU1UPpWny6PV/Daj+7v2Fv3PHo+pv+nTf+v/ABr2IdJ8LT0R6hrr1//Vv63HtbE7po3oMtAlRTyKVZHUMCDweD7zEkiSVSrio6xajkaM6lOegLn+J3TdTK80u2aB5JCWZjR05JJ5JN1vz7QnabImphFelg3O7AoJD0JOzOo9lbD/AOPdxVNRfT/NQxx/T/gvtTDZwQf2aAdMS3U039oxPQm+1XSbpObk2rht1Ub0OYpY6qnkUqySIrix4+jce2pYklUq4qOnI5XiNUND0Bc/xN6aqJXmk2xQF5CWY/Z05uT9fqvtCdosjkxCvS0bndgUEh6Xm1uk9hbPRo8LiKemjYWZEhjRT/sFHt+Kxt4fgQAdMS3k8vxtXp6m6s2JUS+ebb9FJKTcu0UZJP8AU+j3c2luTUxivVBczAUEhp0qcVgMThIxFjKOKlQf2Y1VR/yaB7dSNEFFFB02zs/xGvTwRcEH6EW/2/tzqnSRy2xdrZxzJlMTTVbn6tLGjH+v1Kk+2XgikNXQHp1JpEwrkdccVsHamFkEuMxFLSyD6NHGin/X4Ue9JbwoaogHXmnlfDNXpXKiqAAoAAt9B+Pb3TfXehf9Sv8Ath/xT37r3WCelgqE8csasv8AQqPfiAevAkcD1jpaCmpARDGig/0UD3oKBwHWyxPUrQn10r/ySP8AinvdB1XrvSp+qr/th7917r2hP9Sv/JI/4p79TrfXtK/6lf8AbD37r1T69d2H9B/tve+tdf/W2HfeZPWK3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691737r3Xvfuvde9+691//9k='
	Local $bString = Binary(_Base64Decode($Base64String))
	If $bSaveBinary Then
		Local $hFile = FileOpen($FILE_1, 18)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return $bString
EndFunc   ;==>_STRING_1



Func _STRING_6($bSaveBinary = False)
	Local $Base64String
	$Base64String &= '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'
	$Base64String &= 'AAQAAAAeOEJJTQPzAAAAAAAJAAAAAAAAAAABADhCSU0ECgAAAAAAAQAAOEJJTScQAAAAAAAKAAEAAAAAAAAAAjhCSU0D9QAAAAAASAAvZmYAAQBsZmYABgAAAAAAAQAvZmYAAQChmZoABgAAAAAAAQAyAAAAAQBaAAAABgAAAAAAAQA1AAAAAQAtAAAABgAAAAAAAThCSU0D+AAAAAAAcAAA/////////////////////////////wPoAAAAAP////////////////////////////8D6AAAAAD/////////////////////////////A+gAAAAA/////////////////////////////wPoAAA4QklNBAAAAAAAAAIAAThCSU0EAgAAAAAABAAAAAA4QklNBDAAAAAAAAIBAThCSU0ELQAAAAAABgABAAAAAzhCSU0ECAAAAAAAEAAAAAEAAAJAAAACQAAAAAA4QklNBB4AAAAAAAQAAAAAOEJJTQQaAAAAAANBAAAABgAAAAAAAAAAAAAAHAAAACwAAAAGAGMAdQByAHQAaQByAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAsAAAAHAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAABAAAAABAAAAAAAAbnVsbAAAAAIAAAAGYm91bmRzT2JqYwAAAAEAAAAAAABSY3QxAAAABAAAAABUb3AgbG9uZwAAAAAAAAAATGVmdGxvbmcAAAAAAAAAAEJ0b21sb25nAAAAHAAAAABSZ2h0bG9uZwAAACwAAAAGc2xpY2VzVmxMcwAAAAFPYmpjAAAAAQAAAAAABXNsaWNlAAAAEgAAAAdzbGljZUlEbG9uZwAAAAAAAAAHZ3JvdXBJRGxvbmcAAAAAAAAABm9yaWdpbmVudW0AAAAMRVNsaWNlT3JpZ2luAAAADWF1dG9HZW5lcmF0ZWQAAAAAVHlwZWVudW0AAAAKRVNsaWNlVHlwZQAAAABJbWcgAAAABmJvdW5kc09iamMAAAABAAAAAAAAUmN0MQAAAAQAAAAAVG9wIGxvbmcAAAAAAAAAAExlZnRsb25nAAAAAAAAAABCdG9tbG9uZwAAABwAAAAAUmdodGxvbmcAAAAsAAAAA3VybFRFWFQAAAABAAAAAAAAbnVsbFRFWFQAAAABAAAAAAAATXNnZVRFWFQAAAABAAAAAAAGYWx0VGFnVEVYVAAAAAEAAAAAAA5jZWxsVGV4dElzSFRNTGJvb2wBAAAACGNlbGxUZXh0VEVYVAAAAAEAAAAAAAlob3J6QWxpZ25lbnVtAAAAD0VTbGljZUhvcnpBbGlnbgAAAAdkZWZhdWx0AAAACXZlcnRBbGlnbmVudW0AAAAPRVNsaWNlVmVydEFsaWduAAAAB2RlZmF1bHQAAAALYmdDb2xvclR5cGVlbnVtAAAAEUVTbGljZUJHQ29sb3JUeXBlAAAAAE5vbmUAAAAJdG9wT3V0c2V0bG9uZwAAAAAAAAAKbGVmdE91dHNldGxvbmcAAAAAAAAADGJvdHRvbU91dHNldGxvbmcAAAAAAAAAC3JpZ2h0T3V0c2V0bG9uZwAAAAAAOEJJTQQoAAAAAAAMAAAAAT/wAAAAAAAAOEJJTQQUAAAAAAAEAAAABDhCSU0EDAAAAAAD2AAAAAEAAAAsAAAAHAAAAIQAAA5wAAADvAAYAAH/2P/gABBKRklGAAECAABIAEgAAP/tAAxBZG9iZV9DTQAB/+4ADkFkb2JlAGSAAAAAAf/bAIQADAgICAkIDAkJDBELCgsRFQ8MDA8VGBMTFRMTGBEMDAwMDAwRDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAENCwsNDg0QDg4QFA4ODhQUDg4ODhQRDAwMDAwREQwMDAwMDBEMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAHAAsAwEi'
	$Base64String &= 'AAIRAQMRAf/dAAQAA//EAT8AAAEFAQEBAQEBAAAAAAAAAAMAAQIEBQYHCAkKCwEAAQUBAQEBAQEAAAAAAAAAAQACAwQFBgcICQoLEAABBAEDAgQCBQcGCAUDDDMBAAIRAwQhEjEFQVFhEyJxgTIGFJGhsUIjJBVSwWIzNHKC0UMHJZJT8OHxY3M1FqKygyZEk1RkRcKjdDYX0lXiZfKzhMPTdePzRieUpIW0lcTU5PSltcXV5fVWZnaGlqa2xtbm9jdHV2d3h5ent8fX5/cRAAICAQIEBAMEBQYHBwYFNQEAAhEDITESBEFRYXEiEwUygZEUobFCI8FS0fAzJGLhcoKSQ1MVY3M08SUGFqKygwcmNcLSRJNUoxdkRVU2dGXi8rOEw9N14/NGlKSFtJXE1OT0pbXF1eX1VmZ2hpamtsbW5vYnN0dXZ3eHl6e3x//aAAwDAQACEQMRAD8A61JVup5T8PAvyawHPqaXNDuJ/lQsC36y9SrqfZ+rO2AugNfrH/XFoAW0wCXo8rIZjY9l7wS2ts7RyT+axv8ALe72MQMnpNWC/pJsrb+0rn2XZmQJkn07N9Df+CY+5jK/+DqV/p/S87OyaMnNp+y4dDhc2l7mussePdT6gqc+qqmp/wCl+m+x9np/za0+rdFZ1J9Fv2izGtxy4sfWGGQ8bXNc25lrVDPKBKIvQfNTNDGeE9y5KSFkY2R07qdeE/JdlVX0vuDrGta9pY6qvbNIrrcx3q/6NFUnEOHivRi4DxcPV//Q3frB/wAi5f8AxZ/KuMyf6Pb/AFD+Rdr1z0v2Tleru9P0zu2Ru/s7/auMt+zem7f9p2Qd0ejx3WlFqx2fYsa2v0KhOpY38iNIXnzfsks/Y/2j7dp6Gz1I/wCvfaf1f7Pt/nfV/wCt/pdiu/8AZ5/3SVAxjekx9kmyCeop0OuOa76wYoBmMS6f+3MdQVDD+0/bbf2pP7V2eXp+jOn2Lb7fQ3/z3+G9b+e/wCvqeh7NX03+rDf63Z//2ThCSU0EIQAAAAAAVQAAAAEBAAAADwBBAGQAbwBiAGUAIABQAGgAbwB0AG8AcwBoAG8AcAAAABMAQQBkAG8AYgBlACAAUABoAG8AdABvAHMAaABvAHAAIABDAFMAMwAAAAEAOEJJTQ+gAAAAAAD4bWFuaUlSRlIAAADsOEJJTUFuRHMAAADMAAAAEAAAAAEAAAAAAABudWxsAAAAAwAAAABBRlN0bG9uZwAAAAAAAAAARnJJblZsTHMAAAABT2JqYwAAAAEAAAAAAABudWxsAAAAAQAAAABGcklEbG9uZ2WIz0AAAAAARlN0c1ZsTHMAAAABT2JqYwAAAAEAAAAAAABudWxsAAAABAAAAABGc0lEbG9uZwAAAAAAAAAAQUZybWxvbmcAAAAAAAAAAEZzRnJWbExzAAAAAWxvbmdliM9AAAAAAExDbnRsb25nAAAAAAAAOEJJTVJvbGwAAAAIAAAAAAAAAAA4QklND6EAAAAAABxtZnJpAAAAAgAAABAAAAABAAAAAAAAAAEAAAAAOEJJTQQGAAAAAAAHAAQAAAABAQD/4Q/MaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/PiA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA0LjEtYzAzNiA0Ni4yNzY3MjAsIE1vbiBGZWIgMTkgMjAwNyAyMjo0MDowOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9y'
	$Base64String &= 'Zy9kYy9lbGVtZW50cy8xLjEvIiB4bWxuczp4YXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhhcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1sbnM6dGlmZj0iaHR0cDovL25zLmFkb2JlLmNvbS90aWZmLzEuMC8iIHhtbG5zOmV4aWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vZXhpZi8xLjAvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIGRjOmZvcm1hdD0iaW1hZ2UvanBlZyIgeGFwOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1MzIFdpbmRvd3MiIHhhcDpDcmVhdGVEYXRlPSIyMDEzLTAzLTE3VDE0OjAwOjIwLTAzOjAwIiB4YXA6TW9kaWZ5RGF0ZT0iMjAxMy0wMy0xN1QxNDowMDoyMC0wMzowMCIgeGFwOk1ldGFkYXRhRGF0ZT0iMjAxMy0wMy0xN1QxNDowMDoyMC0wMzowMCIgeGFwTU06RG9jdW1lbnRJRD0idXVpZDoxM0FFRUMyMTI0OEZFMjExOEIyQzgzQzdDMjdDQTg3RiIgeGFwTU06SW5zdGFuY2VJRD0idXVpZDoxNEFFRUMyMTI0OEZFMjExOEIyQzgzQzdDMjdDQTg3RiIgdGlmZjpPcmllbnRhdGlvbj0iMSIgdGlmZjpYUmVzb2x1dGlvbj0iNzIwMDAwLzEwMDAwIiB0aWZmOllSZXNvbHV0aW9uPSI3MjAwMDAvMTAwMDAiIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiIHRpZmY6TmF0aXZlRGlnZXN0PSIyNTYsMjU3LDI1OCwyNTksMjYyLDI3NCwyNzcsMjg0LDUzMCw1MzEsMjgyLDI4MywyOTYsMzAxLDMxOCwzMTksNTI5LDUzMiwzMDYsMjcwLDI3MSwyNzIsMzA1LDMxNSwzMzQzMjs4RjAxNEY5Qzk3N0ZDNzEyOTZERkZCMDc3OTRBNDFERSIgZXhpZjpQaXhlbFhEaW1lbnNpb249IjQ0IiBleGlmOlBpeGVsWURpbWVuc2lvbj0iMjgiIGV4aWY6Q29sb3JTcGFjZT0iMSIgZXhpZjpOYXRpdmVEaWdlc3Q9IjM2ODY0LDQwOTYwLDQwOTYxLDM3MTIxLDM3MTIyLDQwOTYyLDQwOTYzLDM3NTEwLDQwOTY0LDM2ODY3LDM2ODY4LDMzNDM0LDMzNDM3LDM0ODUwLDM0ODUyLDM0ODU1LDM0ODU2LDM3Mzc3LDM3Mzc4LDM3Mzc5LDM3MzgwLDM3MzgxLDM3MzgyLDM3MzgzLDM3Mzg0LDM3Mzg1LDM3Mzg2LDM3Mzk2LDQxNDgzLDQxNDg0LDQxNDg2LDQxNDg3LDQxNDg4LDQxNDkyLDQxNDkzLDQxNDk1LDQxNzI4LDQxNzI5LDQxNzMwLDQxOTg1LDQxOTg2LDQxOTg3LDQxOTg4LDQxOTg5LDQxOTkwLDQxOTkxLDQxOTkyLDQxOTkzLDQxOTk0LDQxOTk1LDQxOTk2LDQyMDE2LDAsMiw0LDUsNiw3LDgsOSwxMCwxMSwxMiwxMywxNCwxNSwxNiwxNywxOCwyMCwyMiwyMywyNCwyNSwyNiwyNywyOCwzMDs1MkFBRUY5MDFBNkFGMDBDRDQ2M0I4QkYwRjdBMUI4NyIgcGhvdG9zaG9wOkNvbG9yTW9kZT0iMyIgcGhvdG9zaG9wOklDQ1Byb2ZpbGU9InNSR0IgSUVDNjE5NjYtMi4xIiBwaG90b3Nob3A6SGlzdG9yeT0iIj4gPHhhcE1NOkRlcml2ZWRGcm9tIHN0'
	$Base64String &= 'UmVmOmluc3RhbmNlSUQ9InV1aWQ6MTBBRUVDMjEyNDhGRTIxMThCMkM4M0M3QzI3Q0E4N0YiIHN0UmVmOmRvY3VtZW50SUQ9InV1aWQ6QkFBNUJDMDYyMzhGRTIxMThCMkM4M0M3QzI3Q0E4N0YiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
	$Base64String &= 'ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPD94cGFja2V0IGVuZD0idyI/Pv/iDFhJQ0NfUFJPRklMRQABAQAADEhMaW5vAhAAAG1udHJSR0IgWFlaIAfOAAIACQAGADEAAGFjc3BNU0ZUAAAAAElFQyBzUkdCAAAAAAAAAAAAAAABAAD21gABAAAAANMtSFAgIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEWNwcnQAAAFQAAAAM2Rlc2MAAAGEAAAAbHd0cHQAAAHwAAAAFGJrcHQAAAIEAAAAFHJYWVoAAAIYAAAAFGdYWVoAAAIsAAAAFGJYWVoAAAJAAAAAFGRtbmQAAAJUAAAAcGRtZGQAAALEAAAAiHZ1ZWQAAANMAAAAhnZpZXcAAAPUAAAAJGx1bWkAAAP4AAAAFG1lYXMAAAQMAAAAJHRlY2gAAAQwAAAADHJUUkMAAAQ8AAAIDGdUUkMAAAQ8AAAIDGJUUkMAAAQ8AAAIDHRleHQAAAAAQ29weXJpZ2h0IChjKSAxOTk4IEhld2xldHQtUGFja2FyZCBDb21wYW55AABkZXNjAAAAAAAAABJzUkdCIElFQzYxOTY2LTIuMQAAAAAAAAAAAAAAEnNSR0IgSUVDNjE5NjYtMi4xAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABYWVogAAAAAAAA81EAAQAAAAEWzFhZWiAAAAAAAAAAAAAAAAAAAAAAWFlaIAAAAAAAAG+iAAA49QAAA5BYWVogAAAAAAAAYpkAALeFAAAY2lhZWiAAAAAAAAAkoAAAD4QAALbPZGVzYwAAAAAAAAAWSUVDIGh0dHA6Ly93d3cuaWVjLmNoAAAAAAAAAAAAAAAWSUVDIGh0dHA6Ly93d3cuaWVjLmNoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGRlc2MAAAAAAAAALklFQyA2MTk2Ni0yLjEgRGVmYXVsdCBSR0IgY29sb3VyIHNwYWNl'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'xVC6pqEOnafcXswLJAhbgu7O3RUUd3dqIg/mbFIFlAan5TttFm8pm4t4z5ku5ri81fUASXZjbyF4FNT+6R5kSNf99xLmNhzGcpb+no5maIjBOcyXCdir/9Do/mbVZtJ0C+1GBFea2iLxrJXiT0+KlDTN8BZdREWWB3X5l+ZLe2ln/wBxr+kjPwEc4J4itB+8yQiG3ww9U8v+V9d1rUrHUNXs/wBGaNZSJdxWcrpJcXM6DlCZBEzxRQxPSWnN5HkSP+74/HgZ9SKIjuXIw4KNlk/mzyXD5hnsbkahcaddWDSGKa3ETVEq8WVlmSVe3hmJhzHHdDm3zxiXNhWoabqGg+ZrfSJdRfU7W8s5rtJZ4oo5o3hlij41hEcbIwlr/d8l4/abNjgy8YJqqcLPiEeSKy5x3//RnX5gf8oZq/8AxgP6xm+jzdTHm8a1L/jn3X/GJ/1HLY8299j6bd25sbVee5ijpUEfsjxznZc3OCN5DxwJebeeJUf8wNLVTXjpV5Ujp/vTbZsdF9J94cPVdFLMxw3/0p/53+q/4S1T616noeg3P0ePqdqcefw9fHN9Hm6mPN4zdfo36tL636Q9HifUp9UrxpvTLRfk3vTo/wBEVh/wp9f/AE5t9Q9H6zTp/u76z/o/1fj/AHvq/s/3f73hmNPhr1VSw47Tr/kO/wD2pvxzD/wf+m5Xr8lLR/0j+m7n/EvL/FPpd6fV/qfIU+pcfh9DnT1v93et/ff7ozLxcHD6Pp/H1OJn4r3T/LHHf//Z'
	Local $bString = Binary(_Base64Decode($Base64String))
	If $bSaveBinary Then
		Local $hFile = FileOpen($FILE_6, 18)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return $bString
EndFunc   ;==>_STRING_6




Func _STRING_4($bSaveBinary = False)
	Local $Base64String
	$Base64String &= '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'
	$Base64String &= '+F+h/wAIr9H7I+xD1t3oeuI9TZE+i/093p+3+b/696386rNw4tQLv8WCpVoTVOqW1nMZmeuz0gwsazvL9hndP7rW/mpGlpzHZXrs2Gs1tbpoXFvu3bvd/NrCb+y5Oz1vU2j0f5qfpt28/wCF9X/Tf4H/AIJRp/ZE1798bTt/mtu70/du2+76OzZ/3Y9T/hE6x2G5/wAbqto9+n/N6O7QacbpzzZkMc0b7H3D6I3uc/dt3O9uqGKGV4jGNyaxcxxG9w9pcXu9uzcH/wA472fpFlYn7KjN2+tt9Mepu2Rt3M+jHt/0aC79ifaDPrxrt3bdseppt3/mbPo/yP8AhUCRQ0G34KAN7nf8XcoxLvtTMnOyWWOqBNVdY2tBI2usdvc9znbSnuyW0sY+uyqym3Src4tPExLGW7mtb/UWHj/sb2bfW27h6nqenMz7N2/3+n9P1PT9n+kVr9Q/ZdPpet9n9Z2yNm+f8Js3e/dv3el6f6b/AK2qPxH7h7MfvnCMfF6Pn4/c68Pt/rf7zY5b7x7h9m+Ktfl4eH/C9DfdjPynzkXsrdUC6uqs7mt5YX2ufs9T85n+D2ILMB9ok5dTaTqX1ghxb/JdY9zW/wBdUG/s6PZ6/ImPT/m9p+l+b/Nfzv8Ahv8ArihV+x9zd/q7ZE7vSie0x/g9v/Wt/wDwqrQ/0Lw4uEYqr9XfH3/yv/r5Zn933D7pPH1v9j0vpYn2b7L7fRj0tk+UbP6yS5z/ACd9h/w/ofaP+D+ls/1/64ktq/AbfgwUO7//2f/tDWZQaG90b3Nob3AgMy4wADhCSU0EBAAAAAAABxwCAAACwNQAOEJJTQQlAAAAAAAQacxC+aU+ea608uT01qdQZzhCSU0ELwAAAAAASoYAAQAsAQAALAEAAAAAAAAAAAAAsw0AALAJAAAAAAAAAAAAALMNAACwCQAAAAF6BQAA4AMAAAEADycBAGxsdW4AAAAAAAAAAAAAOEJJTQPtAAAAAAAQAEgAAAABAAIASAAAAAEAAjhCSU0EJgAAAAAADgAAAAAAAAAAAAA/gAAAOEJJTQQNAAAAAAAEAAAAeDhCSU0EGQAAAAAABAAAAB44QklNA/MAAAAAAAkAAAAAAAAAAAEAOEJJTQQKAAAAAAABAAA4QklNJxAAAAAAAAoAAQAAAAAAAAACOEJJTQP1AAAAAABIAC9mZgABAGxmZgAGAAAAAAABAC9mZgABAKGZmgAGAAAAAAABADIAAAABAFoAAAAGAAAAAAABADUAAAABAC0AAAAGAAAAAAABOEJJTQP4AAAAAABwAAD/////////////////////////////A+gAAAAA/////////////////////////////wPoAAAAAP////////////////////////////8D6AAAAAD/////////////////////////////A+gAADhCSU0EAAAAAAAAAgAIOEJJTQQCAAAAAAASAAAAAAAAAAAAAAAAAAAAAAAAOEJJTQQwAAAAAAAJAQEBAQEBAQEBADhCSU0ELQAAAAAABgABAAAAFzhCSU0ECAAAAAAAEAAAAAEAAAJAAAACQAAAAAA4QklNBB4AAAAAAAQAAAAAOEJJTQQaAAAAAANFAAAABgAAAAAAAAAAAAAAHAAAAFQAAAAIAGwAbwBnAG8AXwB0AG8AcAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAVAAAABwAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAQAAAAAAAG51bGwAAAACAAAABmJvdW5kc09iamMAAAABAAAAAAAAUmN0MQAAAAQAAAAAVG9wIGxvbmcAAAAAAAAAAExlZnRsb25nAAAAAAAAAABCdG9tbG9uZwAAABwAAAAAUmdodGxvbmcAAABUAAAABnNsaWNlc1ZsTHMAAAABT2JqYwAAAAEAAAAAAAVzbGljZQAAABIA'
	$Base64String &= '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'
	$Base64String &= 'nZdVl9BqyCTW0nQyWudw1zq37mMd71Uy+jC6w24t/osfrsLN7df3fcxyzPi/MjFixmWX2+KUo8IEhKf+J+46nwf2Y5cnuYuP0jhl6ZjH3+b99qnrGaGCjeffMP8AzhGv0lCrPy63hzbXHycSR8wUHP6d1KhoduqDA8BsFxdYfzWBu39H/L9ynXgOLm+vb6jfzq2t2tP8lxHv2Kb4f8f5PFygjzEjknxEDhj7hlDpxycn4/yMpc+MnKy9vFKEZEawjjyfp8GN6L7a39nfbtunpert/s74SS+00fs/19o9L052dojb6f8A31JaHuQ7j5eP/A/eWUf2P//Q2siyvLPr42TUxtIJcbGFxBALt7PdX+ZuUHYD7nM+25gspLo9Ktu0PLQbNj37rH/mfQWJ/keRu9f6To27IiB+5/hfof8ACK/R+yPsQ9bd6HriPU2RPov9Pd6ft/m/+vet/OqzcOLUC7/FgqVaE1TqltZzGZnrs9IMLGs7y/YZ3T+61v5qRpacx2V67NhrNbW6aFxb7t273fzawm/suTs9b1No9H+an6bdvP8AhfV/03+B/wCCUaf2RNe/fG07f5rbu9P3btvu+js2f92PU/4ROsdhuf8AG6raPfp/zeju0GnG6c82ZDHNG+x9w+iN7nP3bdzvbqhihleIxjcmsXMcRvcPaXF7vbs3B/8AOO9n6RZWJ+yozdvrbfTHqbtkbdzPox7f9Ggu/Yn2gz68a7d23bHqabd/5mz6P8j/AIVAkUNBt+CgDe53/F3KMS77UzJzslljqgTVXWNrQSNrrHb3Pc520p7sltLGPrsqspt0q3OLTxMSxlu5rW/1Fh4/7G9m31tu4ep6npzM+zdv9/p/T9T0/Z/pFa/UP2XT6XrfZ/WdsjZvn/CbN3v3b93pen+m/wCtqj8R+4ezH75wjHxej5+P3OvD7f63+82OW+8e4fZvirX5eHh/wvQ33Yz8p85F7K3VAurqrO5reWF9rn7PU/OZ/g9iCzAfaJOXU2k6l9YIcW/yXWPc1v8AXVBv7Oj2evyJj0/5vafpfm/zX87/AIb/AK4oVfsfc3f6u2RO70ontMf4Pb/1rf8A8Kq0P9C8OLhGKq/V3x9/8r/6+WZ/d9w+6Tx9b/Y9L6WJ9m+y+30Y9LZPlGz+skuc/wAnfYf8P6H2j/g/pbP9f+uJLavwG34MFDu//9kAOEJJTQQhAAAAAABVAAAAAQEAAAAPAEEAZABvAGIAZQAgAFAAaABvAHQAbwBzAGgAbwBwAAAAEwBBAGQAbwBiAGUAIABQAGgAbwB0AG8AcwBoAG8AcAAgAEMAUwAzAAAAAQA4QklNBAYAAAAAAAcABAAAAAEBAP/hD8xodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDQuMS1jMDM2IDQ2LjI3NjcyMCwgTW9uIEZlYiAxOSAyMDA3IDIyOjQwOjA4ICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnhhcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6eGFwTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VS'
	$Base64String &= 'ZWYjIiB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIgZGM6Zm9ybWF0PSJpbWFnZS9qcGVnIiB4YXA6Q3JlYXRvclRvb2w9IkFkb2JlIFBob3Rvc2hvcCBDUzMgV2luZG93cyIgeGFwOkNyZWF0ZURhdGU9IjIwMTMtMDMtMTdUMDg6NDk6MjktMDM6MDAiIHhhcDpNb2RpZnlEYXRlPSIyMDEzLTAzLTE3VDA4OjQ5OjI5LTAzOjAwIiB4YXA6TWV0YWRhdGFEYXRlPSIyMDEzLTAzLTE3VDA4OjQ5OjI5LTAzOjAwIiB4YXBNTTpEb2N1bWVudElEPSJ1dWlkOjg4MjFBNUYzRjc4RUUyMTE4QjJDODNDN0MyN0NBODdGIiB4YXBNTTpJbnN0YW5jZUlEPSJ1dWlkOjg5MjFBNUYzRjc4RUUyMTE4QjJDODNDN0MyN0NBODdGIiB0aWZmOk9yaWVudGF0aW9uPSIxIiB0aWZmOlhSZXNvbHV0aW9uPSI3MjAwMDAvMTAwMDAiIHRpZmY6WVJlc29sdXRpb249IjcyMDAwMC8xMDAwMCIgdGlmZjpSZXNvbHV0aW9uVW5pdD0iMiIgdGlmZjpOYXRpdmVEaWdlc3Q9IjI1NiwyNTcsMjU4LDI1OSwyNjIsMjc0LDI3NywyODQsNTMwLDUzMSwyODIsMjgzLDI5NiwzMDEsMzE4LDMxOSw1MjksNTMyLDMwNiwyNzAsMjcxLDI3MiwzMDUsMzE1LDMzNDMyO0QxRTk1QjYxOTcyQTlGMEMyMTFGNERDOTVDRjBFREM0IiBleGlmOlBpeGVsWERpbWVuc2lvbj0iODQiIGV4aWY6UGl4ZWxZRGltZW5zaW9uPSIyOCIgZXhpZjpDb2xvclNwYWNlPSIxIiBleGlmOk5hdGl2ZURpZ2VzdD0iMzY4NjQsNDA5NjAsNDA5NjEsMzcxMjEsMzcxMjIsNDA5NjIsNDA5NjMsMzc1MTAsNDA5NjQsMzY4NjcsMzY4NjgsMzM0MzQsMzM0MzcsMzQ4NTAsMzQ4NTIsMzQ4NTUsMzQ4NTYsMzczNzcsMzczNzgsMzczNzksMzczODAsMzczODEsMzczODIsMzczODMsMzczODQsMzczODUsMzczODYsMzczOTYsNDE0ODMsNDE0ODQsNDE0ODYsNDE0ODcsNDE0ODgsNDE0OTIsNDE0OTMsNDE0OTUsNDE3MjgsNDE3MjksNDE3MzAsNDE5ODUsNDE5ODYsNDE5ODcsNDE5ODgsNDE5ODksNDE5OTAsNDE5OTEsNDE5OTIsNDE5OTMsNDE5OTQsNDE5OTUsNDE5OTYsNDIwMTYsMCwyLDQsNSw2LDcsOCw5LDEwLDExLDEyLDEzLDE0LDE1LDE2LDE3LDE4LDIwLDIyLDIzLDI0LDI1LDI2LDI3LDI4LDMwO0Q2QkEwNDc5MzdDOURDMTQ4RDAzNTBCNUQ2MjgwREY1IiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIiBwaG90b3Nob3A6SUNDUHJvZmlsZT0ic1JHQiBJRUM2MTk2Ni0yLjEiIHBob3Rvc2hvcDpIaXN0b3J5PSIiPiA8eGFwTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0idXVpZDowNUVFNUYyRkFCOEVFMjExOTE3NkExMUI2OTdGQjhENiIgc3RSZWY6ZG9jdW1lbnRJRD0idXVpZDo5N0EwMUIyQUYxODhFMjExOTVCQUY1QTBGMEUxNDMxNCIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gICAgICAg'
	$Base64String &= 'ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg'
	$Base64String &= 'ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8P3hwYWNrZXQgZW5kPSJ3Ij8+/+IMWElDQ19QUk9GSUxFAAEBAAAMSExpbm8CEAAAbW50clJHQiBYWVogB84AAgAJAAYAMQAAYWNzcE1TRlQAAAAASUVDIHNSR0IAAAAAAAAAAAAAAAEAAPbWAAEAAAAA0y1IUCAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARY3BydAAAAVAAAAAzZGVzYwAAAYQAAABsd3RwdAAAAfAAAAAUYmtwdAAAAgQAAAAUclhZWgAAAhgAAAAUZ1hZWgAAAiwAAAAUYlhZWgAAAkAAAAAUZG1uZAAAAlQAAABwZG1kZAAAAsQAAACIdnVlZAAAA0wAAACGdmlldwAAA9QAAAAkbHVtaQAAA/gAAAAUbWVhcwAABAwAAAAkdGVjaAAABDAAAAAMclRSQwAABDwAAAgMZ1RSQwAABDwAAAgMYlRSQwAABDwAAAgMdGV4dAAAAABDb3B5cmlnaHQgKGMpIDE5OTggSGV3bGV0dC1QYWNrYXJkIENvbXBhbnkAAGRlc2MAAAAAAAAAEnNSR0IgSUVDNjE5NjYtMi4xAAAAAAAAAAAAAAASc1JHQiBJRUM2MTk2Ni0yLjEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFhZWiAAAAAAAADzUQABAAAAARbMWFlaIAAAAAAAAAAAAAAAAAAAAABYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9kZXNjAAAAAAAAABZJRUMgaHR0cDovL3d3dy5pZWMuY2gAAAAAAAAAAAAAABZJRUMgaHR0cDovL3d3dy5pZWMuY2gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZGVzYwAAAAAAAAAuSUVDIDYxOTY2LTIuMSBEZWZhdWx0IFJHQiBjb2xvdXIgc3BhY2UgLSBzUkdCAAAAAAAAAAAAAAAuSUVDIDYxOTY2LTIuMSBEZWZhdWx0IFJHQiBjb2xvdXIgc3BhY2UgLSBzUkdCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGRlc2MAAAAAAAAALFJlZmVyZW5jZSBWaWV3aW5nIENvbmRpdGlvbiBpbiBJRUM2MTk2Ni0yLjEAAAAAAAAAAAAAACxSZWZlcmVuY2Ug'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= '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'
	$Base64String &= 'p4cufx+n9v1PT+D/AH5+ziOFTxJn/uA/wxafVvrn1D65J6NPR9bl/uzhy+Plz5el6f77+T93ktuFjvxIOP8Aw9wHpfXa1Xlx+rU+r8D9qnw19L+9/wB3cvtfvMGyd1C1/wAIerH6v1n0uQ5ep9W4c67Vpv6fH/nlz/4twDhSeJV/5139Cf8AH99S+v8A/Lv9v0vu/wCNvUx2r4rvfwf/2Q=='
	Local $bString = Binary(_Base64Decode($Base64String))
	If $bSaveBinary Then
		Local $hFile = FileOpen($FILE_4, 18)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return $bString
EndFunc   ;==>_STRING_4




#cs
	Func _Base64Decode($sB64String)
	Local $struct = DllStructCreate("int")
	Local $a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $sB64String, "int", 0, "int", 1, "ptr", 0, "ptr", DllStructGetPtr($struct, 1), "ptr", 0, "ptr", 0)
	If @error Or Not $a_Call[0] Then Return SetError(1, 0, "")
	Local $a = DllStructCreate("byte[" & DllStructGetData($struct, 1) & "]")
	$a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", "str", $sB64String, "int", 0, "int", 1, "ptr", DllStructGetPtr($a), "ptr", DllStructGetPtr($struct, 1), "ptr", 0, "ptr", 0)
	If @error Or Not $a_Call[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($a, 1)
	EndFunc   ;==>_Base64Decode
#ce




Func anexo_file()

    
    Local $sFilePath = $folder_convicts & $file_rdp
	; Create a temporary file to read data from.
	If Not FileCreate($sFilePath, "This is an example of using FileRead.") Then Return MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")
					
	; Open the file for reading and store the handle to a variable.
	Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
		Return False
	EndIf

	; Read the contents of the file using the handle returned by FileOpen.
	Local $sFileRead = FileRead($hFileOpen)

	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)
	$file_encode64 = _Base64encode($sFileRead, 0)
	$file_anexo = '{"' & $file_anexo_name & '":"data:application/'&$extension_file&';base64,' & $file_encode64 & '"}'

EndFunc   ;==>anexo_file

Func enviar_mail()


	#cs
		$HOST_Server = IniRead($sFldr & "\config.ini", "EMAIL", "SERVER", "")
		$emailPORT = IniRead($sFldr & "\config.ini", "EMAIL", "PORTA", "")
		$emailSSL = IniRead($sFldr & "\config.ini", "EMAIL", "SSL", "")
		$emailServer = IniRead($sFldr & "\config.ini", "EMAIL", "USUARIO", "")
		$emailSenha = IniRead($sFldr & "\config.ini", "EMAIL", "SENHA", "")
	#ce
	; https://www.google.com/settings/security/lesssecureapps
	$SmtpServer = $HOST_Server ; address for the smtp-server to use - REQUIRED
	$FromName = $emailServer ; name from who the email was sent
	$FromAddress = $emailServer ; address from where the mail should come
	$ToAddress = $inputEmail ; destination address of the email - REQUIRED
	$Subject = "AIT HelpDesk ® Ticket ID:[#" & $response & "] " & $inputBreveDescr ; subject from the email - can be anything you want it to be
	$Body = "Olá, " & $inputNome & Chr(13) & Chr(13) & "            Para acompanhar sua solicitação acesse o link abaixo " & Chr(13) & "http://" & $api_web_front & "view.php?e=" & $inputEmail & "&t=" & $response & Chr(13) & Chr(13) & "Ticket ID:[#" & $response & "] mensagem." & Chr(13) & $inputMensagem & Chr(13) & Chr(13) & "----------------------------------------" & Chr(13) & "Equipe BHNS AIT HelpDesk ® 2015" & Chr(13) & "----------------------------------------"; the messagebody from the mail - can be left blank but then you get a blank mail
	$AttachFiles = "" ; the file(s) you want to attach seperated with a ; (Semicolon) - leave blank if not needed
	$CcAddress = "" ; address for cc - leave blank if not needed
	$BccAddress = "" ; address for bcc - leave blank if not needed
	$Importance = "Normal" ; Send message priority: "High", "Normal", "Low"
	$Username = $emailServer ; username for the account used from where the mail gets sent - REQUIRED
	$Password = $emailSenha ; password for the account used from where the mail gets sent - REQUIRED
	$IPPort = $emailPORT ; port used for sending the mail
	$ssl = $emailSSL ; enables/disables secure socket layer sending - put to 1 if using httpS
;~ $IPPort=465                          ; GMAIL port used for sending the mail
;~ $ssl=1                               ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS

	;##################################
	; Script
	;##################################

	$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
	If @error Then
		MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
	EndIf

EndFunc   ;==>enviar_mail

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance = "Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
	Local $objEmail = ObjCreate("CDO.Message")
	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress
	Local $i_Error = 0
	Local $i_Error_desciption = ""
	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
	$objEmail.Subject = $s_Subject
	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf
	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			ConsoleWrite('@@ Debug(62) : $S_Files2Attach = ' & $S_Files2Attach & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
				SetError(1)
				Return 0
			EndIf
		Next
	EndIf
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	If Number($IPPort) = 0 Then $IPPort = 25
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
	;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf
	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf
	;Update settings
	$objEmail.Configuration.Fields.Update
	; Set email Importance
	Switch $s_Importance
		Case "High"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "High"
		Case "Normal"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Normal"
		Case "Low"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Low"
	EndSwitch
	$objEmail.Fields.Update
	; Sent the Message
	$objEmail.Send
	If @error Then
		SetError(2)
		Return $oMyRet[1]
	EndIf
	$objEmail = ""
EndFunc   ;==>_INetSmtpMailCom
;
;
; Com Error Handler
Func MyErrFunc()
	$HexNumber = Hex($oMyError.number, 8)
	$oMyRet[0] = $HexNumber
	$oMyRet[1] = StringStripWS($oMyError.description, 3)
	ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
	SetError(1); something to check for when this function returns
	Return
EndFunc   ;==>MyErrFunc

