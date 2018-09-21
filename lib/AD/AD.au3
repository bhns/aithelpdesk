#Tidy_Parameters= /gd 1 /gds 1 /nsdp
#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include-once
#include <Array.au3>
#include <Date.au3>

; #INDEX# =======================================================================================================================
; Title .........: Active Directory Function Library
; AutoIt Version : 3.3.9.2 or later (because of _Date_Time_SystemTimeToDateTimeStr and the latest in COM error handling)
; UDF Version ...: 1.3.0.0
; Language ......: English
; Description ...: A collection of functions for accessing and manipulating Microsoft Active Directory
; Author(s) .....: Jonathan Clelland, water
; Modified.......: 20120824 (YYYMMDD)
; Remarks .......: Please read the ReadMe.txt file for information about installing and using this UDF!
; Contributors ..: feeks, KenE, Sundance, supersonic, Talder, Joe2010, Suba, Ethan Turk, Jerold Schulman, Stephane, card0384
; Resources .....: http://www.wisesoft.co.uk/scripts
;                  http://gallery.technet.microsoft.com/ScriptCenter/en-us
;                  http://www.rlmueller.net/
;                  http://www.activxperts.com/activmonitor/windowsmanagement/scripts/activedirectory/
;
;                  Well known SIDs: http://technet.microsoft.com/en-us/library/cc978401.aspx
;                  AD Schema: http://msdn.microsoft.com/en-us/library/ms675085(VS.85).aspx
;				   Win32 error codes: http://msdn.microsoft.com/en-us/library/ms681381(v=VS.85).aspx
;                  ADSI: http://msdn.microsoft.com/en-us/library/aa772170(v=VS.85).aspx
;                  LDAP: http://msdn.microsoft.com/en-us/library/aa367008(v=VS.85).aspx
;                        http://www.petri.co.il/ldap_search_samples_for_windows_2003_and_exchange.htm
; ===============================================================================================================================

#region #VARIABLES#
; #VARIABLES# ===================================================================================================================
Global $__iAD_Debug = 0 ; Debug level. 0 = no debug information, 1 = Debug info to console, 2 = Debug info to MsgBox, 3 = Debug Info to File
Global $__sAD_DebugFile = @ScriptDir & "\AD_Debug.txt" ; Debug file if $__iAD_Debug is set to 3
Global $__oAD_MyError ; COM Error handler
Global $__oAD_Connection
Global $__oAD_OpenDS
Global $__oAD_RootDSE
Global $__oAD_Command
Global $__oAD_Bind ; Reference to hold the bind cache
Global $__bAD_BindFlags ; Bind flags
Global $sAD_DNSDomain
Global $sAD_HostServer
Global $sAD_Configuration
Global $sAD_UserId = ""
Global $sAD_Password = ""
; ===============================================================================================================================
#endregion #VARIABLES#

#region #CONSTANTS#
; #CONSTANTS# ===================================================================================================================
; ADS_RIGHTS_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx
Global Const $ADS_FULL_RIGHTS = 0xF01FF
Global Const $ADS_USER_UNLOCKRESETACCOUNT = 0x100
Global Const $ADS_OBJECT_READWRITE_ALL = 0x30
; ADS_AUTHENTICATION_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772247(VS.85).aspx
Global Const $ADS_SECURE_AUTH = 0x1
Global Const $ADS_USE_SSL = 0x2
Global Const $ADS_SERVER_BIND = 0x200
; ADS_USER_FLAG_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772300(VS.85).aspx
Global Const $ADS_UF_ACCOUNTDISABLE = 0x2
Global Const $ADS_UF_PASSWD_NOTREQD = 0x20
Global Const $ADS_UF_WORKSTATION_TRUST_ACCOUNT = 0x1000
Global Const $ADS_UF_DONT_EXPIRE_PASSWD = 0x10000
; ADS_GROUP_TYPE_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772263(VS.85).aspx
Global Const $ADS_GROUP_TYPE_GLOBAL_GROUP = 0x2
Global Const $ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 0x4
Global Const $ADS_GROUP_TYPE_UNIVERSAL_GROUP = 0x8
Global Const $ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000
Global Const $ADS_GROUP_TYPE_GLOBAL_SECURITY = BitOR($ADS_GROUP_TYPE_GLOBAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)
Global Const $ADS_GROUP_TYPE_UNIVERSAL_SECURITY = BitOR($ADS_GROUP_TYPE_UNIVERSAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)
Global Const $ADS_GROUP_TYPE_LOCAL_SECURITY = BitOR($ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP, $ADS_GROUP_TYPE_SECURITY_ENABLED)
; ADS_ACETYPE_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772244(VS.85).aspx
Global Const $ADS_ACETYPE_ACCESS_ALLOWED = 0
Global Const $ADS_ACETYPE_ACCESS_DENIED = 0x1
Global Const $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = 0x5
Global Const $ADS_ACETYPE_ACCESS_DENIED_OBJECT = 0x6
; ADS_ACEFLAG_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772242(VS.85).aspx
Global Const $ADS_ACEFLAG_INHERITED_ACE = 0x10
; Global Const $ADS_FLAGTYPE_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772259(VS.85).aspx
Global Const $ADS_FLAG_OBJECT_TYPE_PRESENT = 0x1
; ADS_RIGHTS_ENUM Enumeration. See: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx
Global Const $ADS_RIGHT_DS_SELF = 0x8
Global Const $ADS_RIGHT_DS_WRITE_PROP = 0x20
Global Const $ADS_RIGHT_DS_CONTROL_ACCESS = 0x100
Global Const $ADS_RIGHT_GENERIC_READ = 0x80000000
; GUIDs - LOWER CASE!
Global Const $USER_CHANGE_PASSWORD = "{ab721a53-1e2f-11d0-9819-00aa0040529b}" ; See: http://msdn.microsoft.com/en-us/library/cc223637(PROT.13).aspx
Global Const $SELF_MEMBERSHIP = "{bf9679c0-0de6-11d0-a285-00aa003049e2}" ; See: http://msdn.microsoft.com/en-us/library/cc223513(PROT.10).aspx
Global Const $ALLOWED_TO_AUTHENTICATE = "{68B1D179-0D15-4d4f-AB71-46152E79A7BC}" ; See: http://msdn.microsoft.com/en-us/library/ms684300(VS.85).aspx
Global Const $RECEIVE_AS = "{AB721A56-1E2f-11D0-9819-00AA0040529B}" ; See: http://msdn.microsoft.com/en-us/library/ms684402(VS.85).aspx
Global Const $SEND_AS = "{AB721A54-1E2f-11D0-9819-00AA0040529B}" ; See: http://msdn.microsoft.com/en-us/library/ms684406(VS.85).aspx
Global Const $USER_FORCE_CHANGE_PASSWORD = "{00299570-246D-11D0-A768-00AA006E0529}" ; See: http://msdn.microsoft.com/en-us/library/ms684414(VS.85).aspx
Global Const $USER_ACCOUNT_RESTRICTIONS = "{4C164200-20C0-11D0-A768-00AA006E0529}" ; See: http://msdn.microsoft.com/en-us/library/ms684412(VS.85).aspx
Global Const $VALIDATED_DNS_HOST_NAME = "{72E39547-7B18-11D1-ADEF-00C04FD8D5CD}" ; See: http://msdn.microsoft.com/en-us/library/ms684331(VS.85).aspx
Global Const $VALIDATED_SPN = "{F3A64788-5306-11D1-A9C5-0000F80367C1}" ; See: http://msdn.microsoft.com/en-us/library/ms684417(VS.85).aspx
; ===============================================================================================================================
#endregion #CONSTANTS#

; #CURRENT# =====================================================================================================================
;_AD_Open
;_AD_Close
;_AD_ErrorNotify
;_AD_SamAccountNameToFQDN
;_AD_FQDNToSamAccountName
;_AD_FQDNToDisplayname
;_AD_ObjectExists
;_AD_GetSchemaAttributes
;_AD_GetObjectAttribute
;_AD_IsMemberOf
;_AD_HasFullRights
;_AD_HasUnlockResetRights
;_AD_HasRequiredRights
;_AD_HasGroupUpdateRights
;_AD_GetObjectClass
;_AD_GetUserGroups
;_AD_GetUserPrimaryGroup
;_AD_SetUserPrimaryGroup
;_AD_RecursiveGetMemberOf
;_AD_GetGroupMembers
;_AD_GetGroupMemberOf
;_AD_GetObjectsInOU
;_AD_GetAllOUs
;_AD_ListDomainControllers
;_AD_ListRootDSEAttributes
;_AD_ListRoleOwners
;_AD_GetLastLoginDate
;_AD_IsObjectDisabled
;_AD_IsObjectLocked
;_AD_IsPasswordExpired
;_AD_GetObjectsDisabled
;_AD_GetObjectsLocked
;_AD_GetPasswordExpired
;_AD_GetPasswordDontExpire
;_AD_GetObjectProperties
;_AD_CreateOU
;_AD_CreateUser
;_AD_SetPassword
;_AD_CreateGroup
;_AD_AddUserToGroup
;_AD_RemoveUserFromGroup
;_AD_CreateComputer
;_AD_ModifyAttribute
;_AD_RenameObject
;_AD_MoveObject
;_AD_DeleteObject
;_AD_SetAccountExpire
;_AD_DisablePasswordExpire
;_AD_EnablePasswordExpire
;_AD_EnablePasswordChange
;_AD_DisablePasswordChange
;_AD_UnlockObject
;_AD_DisableObject
;_AD_EnableObject
;_AD_GetPasswordInfo
;_AD_ListExchangeServers
;_AD_ListExchangeMailboxStores
;_AD_GetSystemInfo
;_AD_GetManagedBy
;_AD_GetManager
;_AD_GetGroupAdmins
;_AD_GroupManagerCanModify
;_AD_ListPrintQueues
;_AD_SetGroupManagerCanModify
;_AD_GroupAssignManager
;_AD_GroupRemoveManager
;_AD_AddEmailAddress
;_AD_JoinDomain
;_AD_UnjoinDomain
;_AD_SetPasswordExpire
;_AD_CreateMailbox
;_AD_DeleteMailbox
;_AD_MailEnableGroup
;_AD_IsAccountExpired
;_AD_GetAccountsExpired
;_AD_ListSchemaVersions
;_AD_ObjectExistsInSchema
;_AD_FixSpecialChars
;_AD_GetLastADSIError
;_AD_GetADOProperties
;_AD_SetADOProperties
;_AD_VersionInfo
; ===============================================================================================================================
; #INTERNAL_USE_ONLY#============================================================================================================
;__AD_Int8ToSec
;__AD_LargeInt2Double
;__AD_ObjGet
;__AD_ErrorHandler
;__AD_ReorderACE
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_Open
; Description ...: Opens a connection to Active Directory.
; Syntax.........: _AD_Open([$sAD_UserIdParam = "", $sAD_PasswordParam = ""[, $sAD_DNSDomainParam = "", $sAD_HostServerParam = "", $sAD_ConfigurationParam = ""[, $iAD_Security = 0]]])
; Parameters ....: $sAD_UserIdParam - Optional: UserId credential for authentication. This has to be a valid domain user
;                  $sAD_PasswordParam - Optional: Password for authentication
;                  $sAD_DNSDomainParam - Optional: Active Directory domain name if you want to connect to an alternate domain e.g. DC=microsoft,DC=com
;                  $sAD_HostServerParam - Optional: Name of Domain Controller if you want to connect to a different domain e.g. DC-Server1.microsoft.com
;                  |If you want to connect to a Global Catalog append port 3268 e.g. DC-Server1.microsoft.com:3268
;                  $sAD_ConfigurationParam - Optional: Configuration naming context if you want to connect to a different domain e.g. CN=Configuration,DC=microsoft,DC=com
;                  $iAD_Security - Optional: Specifies the security settings to be used. Can be a combination of the following:
;                  |0: No security settings are used (default)
;                  |1: Sets the connection property "Encrypt Password" to True to encrypt userid and password
;                  |2: The channel is encrypted using Secure Sockets Layer (SSL). AD requires that the Certificate Server be installed to support SSL
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - (No longer used)
;                  |2 - Creation of the COM object to the AD failed. @extended returns error code from ObjCreate
;                  |3 - Open the connection to AD failed. @extended returns error code of the COM error handler.
;                  |    Generated if the User doesn't have query / modify access
;                  |4 - Creation of the RootDSE object failed. @extended returns the error code received by the COM error handler.
;                  |    Generated when connection to the domain isn't successful. @extended returns -2147023541 (0x8007054B)
;                  |5 - Creation of the DS object failed. @extended returns the error code received by the COM error handler
;                  |6 - Parameter $sAD_HostServerParam and $sAD_ConfigurationParam are required when $sAD_DNSDomainParam is specified
;                  |7 - Parameter $sAD_PasswordParam is required when $sAD_UserIdParam is specified
;                  |8 - OpenDSObject method failed. @extended set to error code received from the OpenDSObject method.
;                  |    On Windows XP or lower this shows that $sAD_UserIdParam and/or $sAD_PasswordParam are invalid
;                  |x - For Windows Vista and later: Win32 error code (decimal). To get detailed error information call function _AD_GetLastADSIError
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: To close the connection to the Active Directory, use the _AD_Close function.
;+
;                  _AD_Open will use the alternative credentials $sAD_UserIdParam and $sAD_PasswordParam if passed as parameters.
;                  $sAD_UserIdParam has to be in one of the following forms (assume the samAccountName = DJ)
;                  * Windows Login Name   e.g. "DJ"
;                  * NetBIOS Login Name   e.g. "<DOMAIN>\DJ"
;                  * User Principal Name  e.g. "DJ@domain.com"
;                  All other name formats have NOT been successfully tested (see section "Link").
;+
;                  Connection to an alternate domain (not the domain your computer is a member of) or if your computer is not a domain member
;                  requires $sAD_DNSDomainParam, $sAD_HostServerParam and $sAD_ConfigurationParam as FQDN as well as $sAD_UserIdParam and $sAD_PasswordParam.
;                  Example:
;                  $sAD_DNSDomainParam = "DC=subdomain,DC=example,DC=com"
;                  $sAD_HostServerParam = "servername.subdomain.example.com"
;                  $sAD_ConfigurationParam = "CN=Configuration,DC=subdomain,DC=example,DC=com"
;+
;                  The COM error handler will be initialized only if there doesn't already exist another error handler.
;+
;                  If you specify $sAD_UserIdParam as NetBIOS Login Name or User Principal Name and the OS is Windows Vista or later then _AD_Open will try to
;                  verify the userid/password.
;                  @error will be set to the Win32 error code (decimal). To get detailed error information please call _AD_GetlastADSIError.
;                  For all other OS or if userid is specified as Windows Login Name @error=8.
;                  This is OS dependant because Windows XP doesn't return useful error information.
;                  For Windows Login Name all OS return success even when an error occures. This seems to be caused by secure authentification.
;+
;                  $iAD_Security = 2 activates LDAP/SSL. LDAP/SSL uses port 636 by default.
;                  Note that an SSL server certificate must be configured properly in order to use SSL.
;+
;                  If you want to connect to a specific DC in the current domain the just provide $sAD_HostServerParam and let $sAD_DNSDomainParam and $sAD_ConfigurationParam be blank.
; Related .......: _AD_Close
; Link ..........: http://msdn.microsoft.com/en-us/library/cc223499(PROT.10).aspx (Simple Authentication), http://msdn.microsoft.com/en-us/library/aa746471(VS.85).aspx (ADO)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_Open($sAD_UserIdParam = "", $sAD_PasswordParam = "", $sAD_DNSDomainParam = "", $sAD_HostServerParam = "", $sAD_ConfigurationParam = "", $iAD_Security = 0)

	$__oAD_Connection = ObjCreate("ADODB.Connection") ; Creates a COM object to AD
	If @error Or Not IsObj($__oAD_Connection) Then Return SetError(2, @error, 0)
	; Activate the COM error handler for older AutoIt versions
	If $__iAD_Debug = 0 And Number(StringReplace(@AutoItVersion, ".", "")) < 3392 Then _AD_ErrorNotify(1)
	; ConnectionString Property (ADO): http://msdn.microsoft.com/en-us/library/ms675810.aspx
	$__oAD_Connection.ConnectionString = "Provider=ADsDSOObject" ; Sets Service providertype
	If $sAD_UserIdParam <> "" Then
		If $sAD_PasswordParam = "" Then Return SetError(7, 0, 0)
		$__oAD_Connection.Properties("User ID") = $sAD_UserIdParam ; Authenticate User
		$__oAD_Connection.Properties("Password") = $sAD_PasswordParam ; Authenticate User
		If BitAND($iAD_Security, 1) = 1 Then $__oAD_Connection.Properties("Encrypt Password") = True ; Encrypts userid and password
		$__bAD_BindFlags = $ADS_SERVER_BIND
		If BitAND($iAD_Security, 2) = 2 Then $__bAD_BindFlags = BitOR($__bAD_BindFlags, $ADS_USE_SSL)
		; If userid is the Windows login name then set the flag for secure authentification
		If StringInStr($sAD_UserIdParam, "\") = 0 And StringInStr($sAD_UserIdParam, "@") = 0 Then _
				$__bAD_BindFlags = BitOR($__bAD_BindFlags, $ADS_SECURE_AUTH)
		$__oAD_Connection.Properties("ADSI Flag") = $__bAD_BindFlags
		$sAD_UserId = $sAD_UserIdParam
		$sAD_Password = $sAD_PasswordParam
	EndIf
	; ADO Open Method: http://msdn.microsoft.com/en-us/library/ms676505.aspx
	$__oAD_Connection.Open() ; Open connection to AD
	If @error <> 0 Then Return SetError(3, @error, 0)
	; Connect to another Domain if the Domain parameter is provided
	If $sAD_DNSDomainParam <> "" Then
		If $sAD_HostServerParam = "" Or $sAD_ConfigurationParam = "" Then Return SetError(6, 0, 0)
		$__oAD_RootDSE = ObjGet("LDAP://" & $sAD_HostServerParam & "/RootDSE")
		If @error Or Not IsObj($__oAD_RootDSE) Then Return SetError(4, @error, 0)
		$sAD_DNSDomain = $sAD_DNSDomainParam
		$sAD_HostServer = $sAD_HostServerParam
		$sAD_Configuration = $sAD_ConfigurationParam
	ElseIf $sAD_HostServerParam <> "" Then ;=> allows to connect to a specific DC in the current domain
		$__oAD_RootDSE = ObjGet("LDAP://" & $sAD_HostServerParam & "/RootDSE")
		If @error Or Not IsObj($__oAD_RootDSE) Then Return SetError(4, @error, 0)
		$sAD_DNSDomain = $__oAD_RootDSE.Get("defaultNamingContext")
		$sAD_HostServer = $sAD_HostServerParam
		$sAD_Configuration = $__oAD_RootDSE.Get("ConfigurationNamingContext")
	Else
		$__oAD_RootDSE = ObjGet("LDAP://RootDSE")
		If @error Or Not IsObj($__oAD_RootDSE) Then Return SetError(4, @error, 0)
		$sAD_DNSDomain = $__oAD_RootDSE.Get("defaultNamingContext") ; Retrieve the current AD domain name
		$sAD_HostServer = $__oAD_RootDSE.Get("dnsHostName") ; Retrieve the name of the connected DC
		$sAD_Configuration = $__oAD_RootDSE.Get("ConfigurationNamingContext") ; Retrieve the Configuration naming context
		$__oAD_RootDSE = ObjGet("LDAP://" & $sAD_HostServer & "/RootDSE") ; To guarantee a persistant binding
	EndIf
	; Check userid/password if provided
	If $sAD_UserIdParam <> "" Then
		$__oAD_OpenDS = ObjGet("LDAP:")
		If @error Or Not IsObj($__oAD_OpenDS) Then Return SetError(5, @error, 0)
		$__oAD_Bind = $__oAD_OpenDS.OpenDSObject("LDAP://" & $sAD_HostServer, $sAD_UserIdParam, $sAD_PasswordParam, $__bAD_BindFlags)
		If @error Or Not IsObj($__oAD_Bind) Then ; login error occurred - get extended information
			Local $iAD_Error = @error
			Local $sAD_Hive = "HKLM"
			If @OSArch = "IA64" Or @OSArch = "X64" Then $sAD_Hive = "HKLM64"
			Local $sAD_OSVersion = RegRead($sAD_Hive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentVersion")
			$sAD_OSVersion = StringSplit($sAD_OSVersion, ".")
			If Int($sAD_OSVersion[1]) >= 6 Then ; Delivers detailed error information for Windows Vista and later if debugging is activated
				Local $aAD_Errors = _AD_GetLastADSIError()
				If $aAD_Errors[4] <> 0 Then
					If $__iAD_Debug = 1 Then ConsoleWrite("_AD_Open: " & _ArrayToString($aAD_Errors, @CRLF, 1) & @CRLF)
					If $__iAD_Debug = 2 Then MsgBox(64, "Active Directory Functions - Debug Info - _AD_Open", _ArrayToString($aAD_Errors, @CRLF, 1))
					If $__iAD_Debug = 3 Then FileWrite($__sAD_DebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
							"-------------------" & @CRLF & "_AD_Open: " & _ArrayToString($aAD_Errors, @CRLF, 1) & @CRLF & _
							"========================================================" & @CRLF)
					Return SetError(Dec($aAD_Errors[4]), 0, 0)
				EndIf
				Return SetError(8, $iAD_Error, 0)
			Else
				Return SetError(8, $iAD_Error, 0)
			EndIf
		EndIf
	EndIf
	; ADO Command object as global
	$__oAD_Command = ObjCreate("ADODB.Command")
	$__oAD_Command.ActiveConnection = $__oAD_Connection
	$__oAD_Command.Properties("Page Size") = 1000
	Return 1

EndFunc   ;==>_AD_Open

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_Close
; Description ...: Closes the connection established to Active Directory by _AD_Open.
; Syntax.........: _AD_Close()
; Parameters ....: None
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - Closing the connection to the AD failed. @extended returns the error code received by the COM error handler
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_Open
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_Close()

	$__oAD_Connection.Close() ; Close Connection
	; Reset all Global Variables
	$__iAD_Debug = 0
	$__sAD_DebugFile = @ScriptDir & "\AD_Debug.txt"
	$__oAD_MyError = 0
	$__oAD_Connection = 0
	$sAD_DNSDomain = ""
	$sAD_HostServer = ""
	$sAD_Configuration = ""
	$__oAD_OpenDS = 0
	$__oAD_RootDSE = 0
	$sAD_UserId = ""
	$sAD_Password = ""
	If @error <> 0 Then Return SetError(1, @error, 0) ; Error returned by connection close
	Return 1

EndFunc   ;==>_AD_Close

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ErrorNotify
; Description ...: Set or query the debug level.
; Syntax.........: _AD_ErrorNotify($iAD_Debug[, $sAD_DebugFile = @ScriptDir & "\AD_Debug.txt"])
; Parameters ....: $iAD_Debug - Debug level. Allowed values are:
;                  |-1 - Return the current settings
;                  |0  - Disable debugging
;                  |1  - Enable debugging. Output the debug info to the console
;                  |2  - Enable Debugging. Output the debug info to a MsgBox
;                  |3  - Enable Debugging. Output the debug info to a file defined by $sAD_DebugFile
;                  $sAD_DebugFile - Optional: File to write the debugging info to if $iAD_Debug = 3 (Default = @ScriptDir & "\AD_Debug.txt")
; Return values .: Success (for $iAD_Debug = -1) - one based one-dimensional array with the following elements:
;                  |1 - Debug level. Value from 0 to 3. Check parameter $iAD_Debug for details
;                  |2 - Debug file. File to write the debugging info to as defined by parameter $sAD_DebugFile
;                  |3 - True if the COM error handler has beend set for this UDF. False if debugging is set off or another COM error handler was already stt
;                  Success (for $iAD_Debug = 0) - 1
;                  Success (for $iAD_Debug => 0) - 1, sets @extended to:
;                  |0 - The COM error handler for this UDF was already active
;                  |1 - A COM error handler has successfully been initialized for this UDF
;                  Failure - 0, sets @error to:
;                  |1 - $iAD_Debug is not an integer or < -1 or > 3
;                  |2 - Installation of the custom error handler failed. @extended is set to the error code returned by ObjEvent
;                  |3 - COM error handler already set to another function
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ErrorNotify($iAD_Debug, $sAD_DebugFile = "")

	If Not IsInt($iAD_Debug) Or $iAD_Debug < -1 Or $iAD_Debug > 3 Then Return SetError(1, 0, 0)
	If $sAD_DebugFile = "" Then $sAD_DebugFile = @ScriptDir & "\AD_Debug.txt"
	Switch $iAD_Debug
		Case -1
			Local $avAD_Debug[4] = [3]
			$avAD_Debug[1] = $__iAD_Debug
			$avAD_Debug[2] = $__sAD_DebugFile
			$avAD_Debug[3] = IsObj($__oAD_MyError)
			Return $avAD_Debug
		Case 0
			$__iAD_Debug = 0
			$__sAD_DebugFile = ""
			$__oAD_MyError = 0
		Case Else
			$__iAD_Debug = $iAD_Debug
			$__sAD_DebugFile = $sAD_DebugFile
			; A COM error handler will be initialized only if one does not exist
			If ObjEvent("AutoIt.Error") = "" Then
				$__oAD_MyError = ObjEvent("AutoIt.Error", "__AD_ErrorHandler") ; Creates a custom error handler
				If @error <> 0 Then Return SetError(2, @error, 0)
				Return SetError(0, 1, 1)
			ElseIf ObjEvent("AutoIt.Error") = "__AD_ErrorHandler" Then
				Return SetError(0, 0, 1) ; COM error handler already set by a call to this function
			Else
				Return SetError(3, 0, 0) ; COM error handler already set by another function
			EndIf
	EndSwitch
	Return 1

EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SamAccountNameToFQDN
; Description ...: Returns a Fully Qualified Domain Name (FQDN) from a SamAccountName.
; Syntax.........: _AD_SamAccountNameToFQDN([$sAD_SamAccountName = @UserName])
; Parameters ....: $sAD_SamAccountName - Optional: Security Accounts Manager (SAM) account name (default = @UserName)
; Return values .: Success - Fully Qualified Domain Name (FQDN)
;                  Failure - "", sets @error to:
;                  |1 - No record returned from Active Directory. $sAD_SamAccountName not found
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: A $ sign must be appended to the computer name to generate the FQDN for a sAMAccountName e.g. @ComputerName & "$".
;                  The function escapes the following special characters (# and /). Commas in CN= or OU= have to be escaped by you.
;                  If $sAD_SamAccountName is already a FQDN then the function returns $sAD_SamAccountName unchanged and without raising an error.
; Related .......: _AD_FQDNToSamAccountName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SamAccountNameToFQDN($sAD_SamAccountName = @UserName)

	If StringMid($sAD_SamAccountName, 3, 1) = "=" Then Return $sAD_SamAccountName ; already a FQDN. Return unchanged
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(sAMAccountName=" & $sAD_SamAccountName & ");distinguishedName;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	Local $sAD_FQDN = $oAD_RecordSet.fields(0).value
	Return _AD_FixSpecialChars($sAD_FQDN, 0, "/#")

EndFunc   ;==>_AD_SamAccountNameToFQDN

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_FQDNToSamAccountName
; Description ...: Returns the SamAccountName of a Fully Qualified Domain Name (FQDN).
; Syntax.........: _AD_FQDNToSamAccountName($sAD_FQDN)
; Parameters ....: $sAD_FQDN - Fully Qualified Domain Name (FQDN)
; Return values .: Success - SamAccountName
;                  Failure - "", sets @error to:
;                  |1 - No record returned from Active Directory. $sAD_FQDN not found
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: You have to escape commas in $sAD_FQDN with a backslash. E.g. "CN=Lastname\, Firstname,OU=..."
;                  All other special characters (# and /) are escaped by the function.
;                  If $sAD_FQDN is already a SamAccountName then the function returns $sAD_FQDN unchanged and without raising an error.
; Related .......: _AD_SamAccountNameToFQDN, _AD_FQDNToDisplayname
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_FQDNToSamAccountName($sAD_FQDN)

	If StringMid($sAD_FQDN, 3, 1) <> "=" Then Return $sAD_FQDN ; already a SamAccountName. Return unchanged
	$sAD_FQDN = _AD_FixSpecialChars($sAD_FQDN, 0, "/#") ; Escape special characters in the FQDN
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_FQDN)
	If @error Or Not IsObj($oAD_Object) Or $oAD_Object = 0 Then Return SetError(1, @error, "")
	Local $sAD_Result = $oAD_Object.sAMAccountName
	Return $sAD_Result

EndFunc   ;==>_AD_FQDNToSamAccountName

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_FQDNToDisplayname
; Description ...: Returns the Display Name of a Fully Qualified Domain Name (FQDN) object.
; Syntax.........: _AD_FQDNToDisplayname($sAD_FQDN)
; Parameters ....: $sAD_FQDN - Fully Qualified Domain Name (FQDN)
; Return values .: Success - Display Name
;                  Failure - "", sets @error to:
;                  |x - @error as set by function _AD_GetObjectAttribute
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: You must escape commas in $sAD_FQDN with a backslash. E.g. "CN=Lastname\, Firstname,OU=..."
;                  All other special characters (# and /) are escaped by the function.
;                  The function removes all escape characters (\) from the returned value.
; Related .......: _AD_FQDNToSamAccountName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_FQDNToDisplayname($sAD_FQDN)

	Local $sAD_Name = _AD_GetObjectAttribute($sAD_FQDN, "displayname")
	If @error Then Return SetError(@error, @extended, "")
	Return _AD_FixSpecialChars($sAD_Name, 1)

EndFunc   ;==>_AD_FQDNToDisplayname

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ObjectExists
; Description ...: Returns 1 if exactly one object exists for the given property in the local Active Directory Tree.
; Syntax.........: _AD_ObjectExists([$sAD_Object = @UserName[, $sAD_Property = ""]])
; Parameters ....: $sAD_Object   - Optional: Object (user, computer, group, OU) to check (default = @UserName)
;                  $sAD_Property - Optional: Property to check. If omitted the function tries to determine whether to use sAMAccountname or FQDN
; Return values .: Success - 1, Exactly one object exists for the given property in the local Active Directory Tree
;                  Failure - 0, sets @error to:
;                  |1 - No object found for the specified property
;                  |x - More than one object found for the specified property. x is the number of objects found
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Checking on a computer account requires a "$" (dollar) appended to the sAMAccountName.
;                  To check the existence of an OU use the FQDN of the OU as first parameter because an OU has no SamAccountName.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ObjectExists($sAD_Object = @UserName, $sAD_Property = "")

	If $sAD_Property = "" Then
		$sAD_Property = "samAccountName"
		If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName"
	EndIf
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object, if it exists
	If IsObj($oAD_RecordSet) Then
		If $oAD_RecordSet.RecordCount = 1 Then
			Return 1
		ElseIf $oAD_RecordSet.RecordCount > 1 Then
			Return SetError($oAD_RecordSet.RecordCount, 0, 0)
		Else
			Return SetError(1, 0, 0)
		EndIf
	Else
		Return SetError(1, 0, 0)
	EndIf

EndFunc   ;==>_AD_ObjectExists

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetSchemaAttributes
; Description ...: Enumerates attributes from the AD Schema (those replicated to the Global Catalog, indexed attributes or all).
; Syntax.........: _AD_GetSchemaAttributes([$iAD_Select = 1])
; Parameters ....: $iAD_Select - Optional: Specifies the attributes to be returned:
;                  |1 - Return all attributes (default)
;                  |2 - Return all attributes that are replicated to the Global Catalog
;                  |3 - Return all attributes that are indexed
; Return values .: Success - One-based two dimensional array with the following information for all selected attributes:
;                  |0 - ldapDisplayName of the attribute
;                  |1 - True if the attribute is replicated to the Global Catalog, False or "" of not
;                  |2 - True if the attribute is indexed. Indexed attributes give better performance
;                  Failure - "", sets @error to:
;                  |1 - The LDAP query returned no records or another error occurred
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetSchemaAttributes($iAD_Select = 1)

	Local $aBool[2] = [False, True]
	Local Const $IS_INDEXED = 1
	Local $sAD_Query
	Local $sAD_SchemaNamingContext = $__oAD_RootDSE.Get("SchemaNamingContext")
	If $iAD_Select > 3 Or $iAD_Select < 1 Then $iAD_Select = 1
	If $iAD_Select = 1 Then $sAD_Query = "(objectClass=attributeSchema)" ; all attributes
	If $iAD_Select = 2 Then $sAD_Query = "(&(objectClass=attributeSchema)(isMemberOfPartialAttributeSet=TRUE))" ; attributes replicated to the GC
	If $iAD_Select = 3 Then $sAD_Query = "(&(objectClass=attributeSchema)(searchFlags:1.2.840.113556.1.4.803:=1))" ; indexed attributes
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_SchemaNamingContext & ">;" & $sAD_Query & ";lDAPDisplayName,isMemberOfPartialAttributeSet,searchFlags;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Then Return SetError(1, @error, "")
	Local $aAD_Attributes[$oAD_RecordSet.Recordcount + 1][3] = [[$oAD_RecordSet.Recordcount, 3]]
	Local $iAD_Index = 1
	$oAD_RecordSet.MoveFirst
	While Not $oAD_RecordSet.EOF
		$aAD_Attributes[$iAD_Index][0] = $oAD_RecordSet.Fields("lDAPDisplayName").Value
		$aAD_Attributes[$iAD_Index][1] = $oAD_RecordSet.Fields("isMemberOfPartialAttributeSet").Value
		$aAD_Attributes[$iAD_Index][2] = $aBool[BitAND($oAD_RecordSet.Fields("searchFlags").Value, $IS_INDEXED)]
		$iAD_Index = $iAD_Index + 1
		$oAD_RecordSet.MoveNext
	WEnd
	Return $aAD_Attributes

EndFunc   ;==>_AD_GetSchemaAttributes

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectAttribute
; Description ...: Returns the specified attribute for the named object.
; Syntax.........: _AD_GetObjectAttribute($sAD_Object, $sAD_Attribute)
; Parameters ....: $sAD_Object - sAMAccountName or FQDN of the object the attribute should be retrieved from
;                  $sAD_Attribute - Attribute to be retrieved
; Return values .: Success - Value for the given attribute
;                  Failure - "", sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |2 - $sAD_Attribute does not exist for $sAD_Object. @Extended is set to the error returned by LDAP
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: If the attribute is a single-value the function returns a string otherwise it returns an array.
;                  To get decrypted attributes (GUID, SID, dates etc.) please see _AD_GetObjectProperties.
; Related .......: _AD_ModifyAttribute, _AD_GetObjectProperties
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectAttribute($sAD_Object, $sAD_Attribute)

	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName" ; FQDN provided
	If _AD_ObjectExists($sAD_Object, $sAD_Property) = 0 Then Return SetError(1, 0, "")
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(2, @error, "")
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	Local $sAD_Result = $oAD_Object.Get($sAD_Attribute)
	If @error Then Return SetError(2, @error, "")
	$oAD_Object.PurgePropertyList
	If IsArray($sAD_Result) Then _ArrayInsert($sAD_Result, 0, UBound($sAD_Result, 1))
	Return $sAD_Result

EndFunc   ;==>_AD_GetObjectAttribute

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_IsMemberOf
; Description ...: Returns 1 if the object (user, group, computer) is an immediate member of the group.
; Syntax.........: _AD_IsMemberOf($sAD_Group[, $sAD_Object = @Username[, $bAD_IncludePrimaryGroup = False]])
; Parameters ....: $sAD_Group - Group to be checked for membership. Can be specified as sAMAccountName or Fully Qualified Domain Name (FQDN)
;                  $sAD_Object - Optional: Object type (user, group, computer) to check for membership of $sAD_Group. Can be specified as sAMAccountName or Fully Qualified Domain Name (FQDN) (default = @UserName)
;                  $bAD_IncludePrimaryGroup - Optional: Additionally checks the primary group for object membership (default = False)
; Return values .: Success - 1, Specified object (user, group, computer) is a member of the specified group
;                  Failure - 0, @error set
;                  |0 - $sAD_Object is not a member of $sAD_Group
;                  |1 - $sAD_Group does not exist
;                  |2 - $sAD_Object does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Determines if the object is an immediate member of the group. This function does not verify membership in any nested groups.
; Related .......: _AD_GetUserGroups, _AD_GetUserPrimaryGroup, _AD_RecursiveGetMemberOf
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_IsMemberOf($sAD_Group, $sAD_Object = @UserName, $bAD_IncludePrimaryGroup = False)

	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group)
	Local $iAD_Result = $oAD_Group.IsMember("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	; Check Primary Group if $sAD_Oject isn't a member of the specified group and the flag is set
	If $iAD_Result = 0 And $bAD_IncludePrimaryGroup Then $iAD_Result = (_AD_GetUserPrimaryGroup($sAD_Object) = $sAD_Group)
	; Abs is necessary to make it work for AutoIt versions < 3.3.2.0 with bug #1068
	Return Abs($iAD_Result)

EndFunc   ;==>_AD_IsMemberOf

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_HasFullRights
; Description ...: Returns 1 if the given user has full rights over the given group or user.
; Syntax.........: _AD_HasFullRights($sAD_Object[, $sAD_User = @UserName])
; Parameters ....: $sAD_Object - Group or User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $sAD_User - Optional: User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or SamAccountName (default = @UserName)
; Return values .: Success - 1, Specified user has full rights over the given group or user
;                  Failure - 0, @error set
;                  |0 - $sAD_User does not have full rights over $sAD_Object
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_Object does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_HasUnlockResetRights, _AD_HasRequiredRights, _AD_HasGroupUpdateRights
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx (ADS_RIGHTS_ENUM Enumeration)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_HasFullRights($sAD_Object, $sAD_User = @UserName)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $aAD_MemberOf, $aAD_TrusteeArray, $sAD_TrusteeGroup
	$aAD_MemberOf = _AD_GetUserGroups($sAD_User, 1)
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	If IsObj($oAD_Object) Then
		Local $oAD_Security = $oAD_Object.Get("ntSecurityDescriptor")
		Local $oAD_DACL = $oAD_Security.DiscretionaryAcl
		For $oAD_ACE In $oAD_DACL
			$aAD_TrusteeArray = StringSplit($oAD_ACE.Trustee, "\")
			$sAD_TrusteeGroup = $aAD_TrusteeArray[$aAD_TrusteeArray[0]]
			For $iCount1 = 0 To UBound($aAD_MemberOf) - 1
				If StringInStr($aAD_MemberOf[$iCount1], "CN=" & $sAD_TrusteeGroup & ",") And _
						$oAD_ACE.AccessMask = $ADS_FULL_RIGHTS Then Return 1
			Next
		Next
	EndIf
	Return 0

EndFunc   ;==>_AD_HasFullRights

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_HasUnlockResetRights
; Description ...: Returns 1 if the given user has unlock and password reset rights on the object.
; Syntax.........: _AD_HasUnlockResetRights($sAD_Object[, $sAD_User = @UserName])
; Parameters ....: $sAD_Object - Group or User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $sAD_User - Optional: User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or SamAccountName (default = @UserName)
; Return values .: Success - 1, Specified user has unlock and password reset rights over the given group or user
;                  Failure - 0, @error set
;                  |0 - $sAD_User does not have unlock and password reset rights over $sAD_Object
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_Object does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_HasFullRights, _AD_HasRequiredRights, _AD_HasGroupUpdateRights
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx (ADS_RIGHTS_ENUM Enumeration)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_HasUnlockResetRights($sAD_Object, $sAD_User = @UserName)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $aAD_MemberOf, $aAD_TrusteeArray, $sAD_TrusteeGroup
	$aAD_MemberOf = _AD_GetUserGroups($sAD_User, 1)
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	If IsObj($oAD_Object) Then
		Local $oAD_Security = $oAD_Object.Get("ntSecurityDescriptor")
		Local $oAD_DACL = $oAD_Security.DiscretionaryAcl
		For $oAD_ACE In $oAD_DACL
			$aAD_TrusteeArray = StringSplit($oAD_ACE.Trustee, "\")
			$sAD_TrusteeGroup = $aAD_TrusteeArray[$aAD_TrusteeArray[0]]
			For $iCount1 = 0 To UBound($aAD_MemberOf) - 1
				If StringInStr($aAD_MemberOf[$iCount1], "CN=" & $sAD_TrusteeGroup & ",") And _
						BitAND($oAD_ACE.AccessMask, $ADS_USER_UNLOCKRESETACCOUNT) = $ADS_USER_UNLOCKRESETACCOUNT Then Return 1
			Next
		Next
	EndIf
	Return 0

EndFunc   ;==>_AD_HasUnlockResetRights

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_HasRequiredRights
; Description ...: Returns 1 if the given user has the required rights on the object.
; Syntax.........: _AD_HasRequiredRights($sAD_Object[, $iAD_Right = 983551[, $sAD_User = @UserName]])
; Parameters ....: $sAD_Object - Group or User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $sAD_Right - Optional: Access mask constant to be checked (default = 983551 (Full rights)).
;                  |Full rights is the combination of the following rights:
;                  |ADS_RIGHT_DELETE                   = 0x10000
;                  |ADS_RIGHT_READ_CONTROL             = 0x20000
;                  |ADS_RIGHT_WRITE_DAC                = 0x40000
;                  |ADS_RIGHT_WRITE_OWNER              = 0x80000
;                  |ADS_RIGHT_DS_CREATE_CHILD          = 0x1
;                  |ADS_RIGHT_DS_DELETE_CHILD          = 0x2
;                  |ADS_RIGHT_ACTRL_DS_LIST            = 0x4
;                  |ADS_RIGHT_DS_SELF                  = 0x8
;                  |ADS_RIGHT_DS_READ_PROP             = 0x10
;                  |ADS_RIGHT_DS_WRITE_PROP            = 0x20
;                  |ADS_RIGHT_DS_DELETE_TREE           = 0x40
;                  |ADS_RIGHT_DS_LIST_OBJECT           = 0x80
;                  |ADS_RIGHT_DS_CONTROL_ACCESS        = 0x100
;                  $sAD_User - Optional: User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or SamAccountName (default = @UserName)
; Return values .: Success - 1, Specified user has the required rights over the given group or user
;                  Failure - 0, @error set
;                  |0 - $sAD_User does not have the required rights over $sAD_Object
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_Object does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_HasFullRights, _AD_HasUnlockResetRights, _AD_HasGroupUpdateRights
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx (ADS_RIGHTS_ENUM Enumeration)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_HasRequiredRights($sAD_Object, $iAD_Right = 983551, $sAD_User = @UserName)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $aAD_MemberOf, $aAD_TrusteeArray, $sAD_TrusteeGroup
	$aAD_MemberOf = _AD_GetUserGroups($sAD_User, 1)
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	If IsObj($oAD_Object) Then
		Local $oAD_Security = $oAD_Object.Get("ntSecurityDescriptor")
		Local $oAD_DACL = $oAD_Security.DiscretionaryAcl
		For $oAD_ACE In $oAD_DACL
			$aAD_TrusteeArray = StringSplit($oAD_ACE.Trustee, "\")
			$sAD_TrusteeGroup = $aAD_TrusteeArray[$aAD_TrusteeArray[0]]
			For $iCount1 = 0 To UBound($aAD_MemberOf) - 1
				If StringInStr($aAD_MemberOf[$iCount1], "CN=" & $sAD_TrusteeGroup & ",") And _
						BitAND($oAD_ACE.AccessMask, $iAD_Right) = $iAD_Right Then Return 1
			Next
		Next
	EndIf
	Return 0

EndFunc   ;==>_AD_HasRequiredRights

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_HasGroupUpdateRights
; Description ...: Returns 1 if the given user has rights to update the group membership of the object.
; Syntax.........: _AD_HasGroupUpdateRights($sAD_Object[, $sAD_User = @UserName])
; Parameters ....: $sAD_Object - Group to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $sAD_User - Optional: User to be checked. Can be specified as Fully Qualified Domain Name (FQDN) or SamAccountName (default = @UserName)
; Return values .: Success - 1, Specified user has the rights to update the group membership on the given group
;                  Failure - 0, @error set
;                  |0 - $sAD_User does not have the rights to update the group membership on $sAD_Object
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_Object does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_HasFullRights, _AD_HasUnlockResetRights, _AD_HasRequiredRights
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772285(VS.85).aspx (ADS_RIGHTS_ENUM Enumeration)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_HasGroupUpdateRights($sAD_Object, $sAD_User = @UserName)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $aAD_MemberOf, $aAD_TrusteeArray, $sAD_TrusteeGroup
	$aAD_MemberOf = _AD_GetUserGroups($sAD_User, 1)
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	If IsObj($oAD_Object) Then
		Local $oAD_Security = $oAD_Object.Get("ntSecurityDescriptor")
		Local $oAD_DACL = $oAD_Security.DiscretionaryAcl
		For $oAD_ACE In $oAD_DACL
			$aAD_TrusteeArray = StringSplit($oAD_ACE.Trustee, "\")
			$sAD_TrusteeGroup = $aAD_TrusteeArray[$aAD_TrusteeArray[0]]
			For $iCount1 = 0 To UBound($aAD_MemberOf) - 1
				If StringInStr($aAD_MemberOf[$iCount1], "CN=" & $sAD_TrusteeGroup & ",") And _
						BitAND($oAD_ACE.AccessMask, $ADS_OBJECT_READWRITE_ALL) = $ADS_OBJECT_READWRITE_ALL Then Return 1
			Next
		Next
	EndIf
	Return 0

EndFunc   ;==>_AD_HasGroupUpdateRights

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectClass
; Description ...: Returns the main class (also called structural class) of an object ("user", "group" etc.).
; Syntax.........: _AD_GetObjectClass($sAD_Object[, $bAD_All = False])
; Parameters ....: $sAD_Object - Object for which the main class should be returned. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $bAD_All    - Optional: Returns the main class plus the superior classes from which the main class is deduced hierarchically (default = False)
; Return values .: Success - Main class of the specified object if $bAD_All = False or an zero-based array of the main plus the superior classes if $bAD_All = True
;                  Failure - "", sets @error to:
;                  |1 - Specified object does not exist
;                  |2 - The LDAP query returned no record. @extended is set to the error returned by LDAP
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectClass($sAD_Object, $bAD_All = False)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, "")
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName" ; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	If @error Or Not IsObj($oAD_RecordSet) Then Return SetError(2, @error, "")
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	If $bAD_All Then Return $oAD_Object.ObjectClass
	Return $oAD_Object.Class

EndFunc   ;==>_AD_GetObjectClass

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetUserGroups
; Description ...: Returns an array of group names that the user is immediately a member of.
; Syntax.........: _AD_GetUserGroups([$sAD_User = @UserName[, $bAD_IncludePrimaryGroup = False]])
; Parameters ....: $sAD_User                - Optional: User for which the group membership is to be returned (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $bAD_IncludePrimaryGroup - Optional: include the primary group to the returned list (default = False)
; Return values .: Success - Returns an one-based one dimensional array of group names (FQDN) the user is a member of
;                  Failure - "", sets @error to:
;                  |1 - Specified user does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Works for computers or groups as well.
; Related .......: _AD_IsMemberOf, _AD_GetUserPrimaryGroup, _AD_RecursiveGetMemberOf
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetUserGroups($sAD_User = @UserName, $bAD_IncludePrimaryGroup = False)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, "")
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_Property = "distinguishedName" ; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_User & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the FQDN for the logged on user
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the logged on user
	Local $aAD_Groups = $oAD_Object.GetEx("memberof")
	If IsArray($aAD_Groups) Then
		If $bAD_IncludePrimaryGroup Then _ArrayAdd($aAD_Groups, _AD_GetUserPrimaryGroup($sAD_User))
		_ArrayInsert($aAD_Groups, 0, UBound($aAD_Groups))
	Else
		Local $aAD_Groups[1] = [0]
		If $bAD_IncludePrimaryGroup Then _ArrayAdd($aAD_Groups, _AD_GetUserPrimaryGroup($sAD_User))
		$aAD_Groups[0] = UBound($aAD_Groups) - 1
	EndIf
	Return $aAD_Groups

EndFunc   ;==>_AD_GetUserGroups

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetUserPrimaryGroup
; Description ...: Returns the primary group the user is assigned to.
; Syntax.........: _AD_GetUserPrimaryGroup([$sAD_User = @UserName])
; Parameters ....: $sAD_User - Optional: User for which the primary group is to be returned (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - Primary group (FQDN) the user is assigned to.
;                  Failure - "", sets @error to:
;                  |1 - Specified user does not exist
;                  |2 - A primary group couldn't be found for the specified user
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsMemberOf, _AD_GetUserGroups, _AD_RecursiveGetMemberOf, _AD_SetUserPrimaryGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetUserPrimaryGroup($sAD_User = @UserName)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, "")
	Local $sAD_Property = "samAccountName"
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_Property = "distinguishedName"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_User & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the FQDN for the logged on user
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the logged on user
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(objectCategory=group);cn,primaryGroupToken,DistinguishedName;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute
	While Not $oAD_RecordSet.EOF
		If $oAD_RecordSet.Fields("primaryGroupToken").Value = $oAD_Object.primaryGroupID Then _
				Return $oAD_RecordSet.Fields("DistinguishedName").Value
		$oAD_RecordSet.MoveNext
	WEnd
	Return SetError(2, 0, "")

EndFunc   ;==>_AD_GetUserPrimaryGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SetUserPrimaryGroup
; Description ...: Sets the users primary group.
; Syntax.........: _AD_SetUserPrimaryGroup($sAD_User, $sAD_Group)
; Parameters ....: $sAD_User - User for which the primary group is to be set. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $sAD_Group - Group to be set as the primary group for the specified user. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_Group does not exist
;                  |3 - $sAD_User must be a member of $sAD_Group
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Talder
; Modified.......:
; Remarks .......:
; Related .......: _AD_AddUserToGroup, _AD_GetUserPrimaryGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetUserPrimaryGroup($sAD_User, $sAD_Group)

	If Not _AD_ObjectExists($sAD_User) Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_Group) Then Return SetError(2, 0, 0)
	If Not _AD_IsMemberOf($sAD_Group, $sAD_User) Then Return SetError(3, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMACccountName provided
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMACccountName provided
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_User) ; Retrieve the COM Object for the user
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group) ; Retrieve the COM Object for the group
	$oAD_Group.GetInfoEx(_ArrayCreate("primaryGroupToken"), 0)
	$oAD_User.primaryGroupID = $oAD_Group.primaryGroupToken
	$oAD_User.SetInfo()
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_SetUserPrimaryGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_RecursiveGetMemberOf
; Description ...: Takes a group, user or computer and recursively returns a list of groups the object is a member of.
; Syntax.........: _AD_RecursiveGetMemberOf($sAD_Object[, $iAD_Depth = 10[, $bAD_ListInherited = True[, $bAD_FQDN = True]]])
; Parameters ....: $sAD_Object - User, group or computer for which the group membership is to be returned. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
;                  $iAD_Depth - Optional: Maximum depth of recursion (default = 10)
;                  $bAD_ListInherited - Optional: Defines if the function returns the group(s) it was inherited from (default = True)
;                  $bAD_FQDN - Optional: Specifies the attribute to be returned. True = distinguishedName (FQDN), False = SamAccountName (default = True)
; Return values .: Success - Returns an one-based one dimensional array of group names (FQDN or sAMAccountName) the user or group is a member of
;                  Failure - "", sets @error to:
;                  |1 - Specified user, group or computer does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This function traverses the groups that the object is immediately a member of while also checking its group membership.
;                  For groups that are inherited, the return is the FQDN or sAMAccountname of the group, user or computer, and the FQDN(s) or sAMAccountname(s) of the group(s) it
;                  was inherited from, seperated by '|'(s) if flag $bAD_ListInherited is set to True.
;+
;                  If flag $bAD_ListInherited is set to False then the group names are sorted and only unique groups are returned.
; Related .......: _AD_IsMemberOf, _AD_GetUserGroups, _AD_GetUserPrimaryGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_RecursiveGetMemberOf($sAD_Object, $iAD_Depth = 10, $bAD_ListInherited = True, $bAD_FQDN = True)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, "")
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $iCount1, $iCount2
	Local $sAD_Field = "distinguishedName"
	If Not $bAD_FQDN Then $sAD_Field = "samaccountname"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(member=" & $sAD_Object & ");" & $sAD_Field & ";subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	Local $aAD_Groups[$oAD_RecordSet.RecordCount + 1] = [0]
	If $oAD_RecordSet.RecordCount = 0 Then Return $aAD_Groups
	$oAD_RecordSet.MoveFirst
	$iCount1 = 1
	Local $aAD_TempMemberOf[1]
	Do
		$aAD_Groups[$iCount1] = $oAD_RecordSet.Fields(0).Value
		If $iAD_Depth > 0 Then
			$aAD_TempMemberOf = _AD_RecursiveGetMemberOf($aAD_Groups[$iCount1], $iAD_Depth - 1, $bAD_ListInherited, $bAD_FQDN)
			If $bAD_ListInherited Then
				For $iCount2 = 1 To $aAD_TempMemberOf[0]
					$aAD_TempMemberOf[$iCount2] &= "|" & $aAD_Groups[$iCount1]
				Next
			EndIf
			_ArrayDelete($aAD_TempMemberOf, 0)
			_ArrayConcatenate($aAD_Groups, $aAD_TempMemberOf)
		EndIf
		$iCount1 += 1
		$oAD_RecordSet.MoveNext
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	If $bAD_ListInherited = False Then
		_ArraySort($aAD_Groups, 0, 1)
		$aAD_Groups = _ArrayUnique($aAD_Groups, 1, 1)
	EndIf
	$aAD_Groups[0] = UBound($aAD_Groups) - 1
	Return $aAD_Groups

EndFunc   ;==>_AD_RecursiveGetMemberOf

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetGroupMembers
; Description ...: Returns an array of group members.
; Syntax.........: _AD_GetGroupMembers($sAD_FQDN)
; Parameters ....: $sAD_Group - Group to retrieve members from. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - Returns an one-based one dimensional array of names (FQDN) that are members of the specified group
;                  Failure - "", sets @error to:
;                  |1 - Specified group does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: If the group has no members, _AD_GetGroupMembers returns an array with one element (row count) set to 0
; Related .......: _AD_GetGroupMemberOf
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetGroupMembers($sAD_Group)

	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, "")
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	Local $sAD_Range, $iAD_RangeModifier, $oAD_RecordSet
	Local $aAD_Members[1]
	Local $iCount1 = 0
	Local $aAD_Membersadd
	While 1
		$iAD_RangeModifier = $iCount1 * 1000
		$sAD_Range = "Range=" & $iAD_RangeModifier & "-" & $iAD_RangeModifier + 999
		$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Group & ">;;member;" & $sAD_Range & ";base"
		$oAD_RecordSet = $__oAD_Command.Execute
		$aAD_Membersadd = $oAD_RecordSet.fields(0).Value
		If $aAD_Membersadd = 0 Then ExitLoop
		ReDim $aAD_Members[UBound($aAD_Members) + 1000]
		For $iCount2 = $iAD_RangeModifier + 1 To $iAD_RangeModifier + 1000
			$aAD_Members[$iCount2] = $aAD_Membersadd[$iCount2 - $iAD_RangeModifier - 1]
		Next
		$iCount1 += 1
		$oAD_RecordSet.Close
		$oAD_RecordSet = 0
	WEnd
	$iAD_RangeModifier = $iCount1 * 1000
	$sAD_Range = "Range=" & $iAD_RangeModifier & "-*"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Group & ">;;member;" & $sAD_Range & ";base"
	$oAD_RecordSet = $__oAD_Command.Execute
	$aAD_Membersadd = $oAD_RecordSet.fields(0).Value
	ReDim $aAD_Members[UBound($aAD_Members) + UBound($aAD_Membersadd)]
	For $iCount2 = $iAD_RangeModifier + 1 To $iAD_RangeModifier + UBound($aAD_Membersadd)
		$aAD_Members[$iCount2] = $aAD_Membersadd[$iCount2 - $iAD_RangeModifier - 1]
	Next
	$oAD_RecordSet.Close
	$aAD_Members[0] = UBound($aAD_Members) - 1
	Return $aAD_Members

EndFunc   ;==>_AD_GetGroupMembers

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetGroupMemberOf
; Description ...: Returns an array of group membership.
; Syntax.........: _AD_GetGroupMemberOf($sAD_Group)
; Parameters ....: $sAD_Group - Group for which membership in other groups is to be retrieved. Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - Returns an one-based one dimensional array of group names (FQDN) that the specified group is a member of
;                  Failure - "", sets @error to:
;                  |1 - Specified group does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_GetGroupMembers
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetGroupMemberOf($sAD_Group)

	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, "")
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	Local $iAD_RangeModifier, $sAD_Range, $oAD_RecordSet, $aAD_Membersadd
	Local $aAD_MemberOf[1]
	Local $iCount1 = 0
	While 1
		$iAD_RangeModifier = $iCount1 * 1000
		$sAD_Range = "Range=" & $iAD_RangeModifier & "-" & $iAD_RangeModifier + 999
		$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Group & ">;;memberof;" & $sAD_Range & ";base"
		$oAD_RecordSet = $__oAD_Command.Execute
		$aAD_Membersadd = $oAD_RecordSet.fields(0).Value
		If $aAD_Membersadd = 0 Then ExitLoop
		ReDim $aAD_MemberOf[UBound($aAD_MemberOf) + 1000]
		For $iCount2 = $iAD_RangeModifier + 1 To $iAD_RangeModifier + 1000
			$aAD_MemberOf[$iCount2] = $aAD_Membersadd[$iCount2 - $iAD_RangeModifier - 1]
		Next
		$iCount1 += 1
		$oAD_RecordSet.Close
	WEnd
	$iAD_RangeModifier = $iCount1 * 1000
	$sAD_Range = "Range=" & $iAD_RangeModifier & "-*"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Group & ">;;memberof;" & $sAD_Range & ";base"
	$oAD_RecordSet = $__oAD_Command.Execute
	$aAD_Membersadd = $oAD_RecordSet.fields(0).Value
	ReDim $aAD_MemberOf[UBound($aAD_MemberOf) + UBound($aAD_Membersadd)]
	For $iCount2 = $iAD_RangeModifier + 1 To $iAD_RangeModifier + UBound($aAD_Membersadd)
		$aAD_MemberOf[$iCount2] = $aAD_Membersadd[$iCount2 - $iAD_RangeModifier - 1]
	Next
	$oAD_RecordSet.Close
	$aAD_MemberOf[0] = UBound($aAD_MemberOf) - 1
	Return $aAD_MemberOf

EndFunc   ;==>_AD_GetGroupMemberOf

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectsInOU
; Description ...: Returns a filtered array of objects and attributes for a given OU or just the number of records if $bAD_Count is True.
; Syntax.........: _AD_GetObjectsInOU($sAD_OU[, $sAD_Filter = "(name=*)"[, $iAD_SearchScope = 2[, $sAD_DataToRetrieve =  "sAMAccountName"[, $sAD_SortBy = "sAMAccountName"[, $bAD_Count = False]]]]])
; Parameters ....: $sAD_OU - The OU to retrieve from (FQDN) (default = "", equals "search the whole AD tree")
;                  $sAD_Filter - Optional: An additional LDAP filter if required (default = "(name=*)")
;                  $iAD_SearchScope - Optional: 0 = base, 1 = one-level, 2 = sub-tree (default)
;                  $sAD_DataToRetrieve - Optional: A comma-seperated list of attributes to retrieve (default = "sAMAccountName").
;                  |More than one attribute will create a 2-dimensional array
;                  $sAD_SortBy - Optional: name of the attribute the resulting array will be sorted upon (default = "sAMAccountName").
;                  |To completely suppress sorting (even the default sort) set this parameter to "". This improves performance when doing large queries
;                  $bAD_Count - Optional: If set to True only returns the number of records returned by the query (default = "False")
; Return values .: Success - Number of records retrieved or a one or two dimensional array of objects and attributes in the given OU. First entry is for the given OU itself
;                  Failure - "", sets @error to:
;                  |1 - Specified OU does not exist
;                  |2 - No records returned from Active Directory. $sAD_DataToRetrieve is invalid (attribute may not exist). @extended is set to the error returned by LDAP
;                  |3 - No records returned from Active Directory. $sAD_Filter didn't return a record
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Multi-value attributes are returned as string with the pipe character (|) as separator.
;+
;                  The default filter returns an array including one record for the OU itself. To exclude the OU use a different filter that doesn't include the OU
;                  e.g. "(&(objectcategory=person)(objectclass=user)(name=*))"
;+
;                  To make sure that all properties you specify in $sAD_DataToRetrieve exist in the AD you can use _AD_ObjectExistsInSchema.
;+
;                  The following examples illustrate the use of the escaping mechanism in the LDAP filter:
;                    (o=Parens R Us \28for all your parenthetical needs\29)
;                    (cn=*\2A*)
;                    (filename=C:\5cMyFile)
;                    (bin=\00\00\00\04)
;                    (sn=Lu\c4\8di\c4\87)
;                  The first example shows the use of the escaping mechanism to represent parenthesis characters.
;                  The second shows how to represent a "*" in a value, preventing it from being interpreted as a substring indicator.
;                  The third illustrates the escaping of the backslash character.
;                  The fourth example shows a filter searching for the four-byte value 0x00000004, illustrating the use of the escaping mechanism to
;                  represent arbitrary data, including NUL characters.
;                  The final example illustrates the use of the escaping mechanism to represent various non-ASCII UTF-8 characters.
; Related .......: _AD_GetAllOUs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectsInOU($sAD_OU = "", $sAD_Filter = "(name=*)", $iAD_SearchScope = 2, $sAD_DataToRetrieve = "sAMAccountName", $sAD_SortBy = "sAMAccountName", $bAD_Count = False)

	If $sAD_OU = "" Then
		$sAD_OU = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_OU, "distinguishedName") = 0 Then Return SetError(1, 0, "")
	EndIf
	Local $iCount2, $aAD_DataToRetrieve, $aTemp
	If $sAD_DataToRetrieve = "" Then $sAD_DataToRetrieve = "sAMAccountName"
	$sAD_DataToRetrieve = StringStripWS($sAD_DataToRetrieve, 8)
	$__oAD_Command.Properties("Searchscope") = $iAD_SearchScope
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_OU & ">;" & $sAD_Filter & ";" & $sAD_DataToRetrieve
	$__oAD_Command.Properties("Sort On") = $sAD_SortBy
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Then Return SetError(2, @error, "")
	Local $iCount1 = $oAD_RecordSet.RecordCount
	If $iCount1 = 0 Then
		If $bAD_Count Then Return SetError(3, 0, 0)
		Return SetError(3, 0, "")
	EndIf
	If $bAD_Count Then Return $iCount1
	If StringInStr($sAD_DataToRetrieve, ",") Then
		$aAD_DataToRetrieve = StringSplit($sAD_DataToRetrieve, ",")
		Local $aAD_Objects[$iCount1 + 1][$aAD_DataToRetrieve[0]]
		$aAD_Objects[0][0] = $iCount1
		$aAD_Objects[0][1] = $aAD_DataToRetrieve[0]
		$iCount2 = 1
		$oAD_RecordSet.MoveFirst
		Do
			For $iCount1 = 1 To $aAD_DataToRetrieve[0]
				If IsArray($oAD_RecordSet.Fields($aAD_DataToRetrieve[$iCount1]).Value) Then
					$aTemp = $oAD_RecordSet.Fields($aAD_DataToRetrieve[$iCount1]).Value
					$aAD_Objects[$iCount2][$iCount1 - 1] = _ArrayToString($aTemp)
				Else
					$aAD_Objects[$iCount2][$iCount1 - 1] = $oAD_RecordSet.Fields($aAD_DataToRetrieve[$iCount1]).Value
				EndIf
			Next
			$oAD_RecordSet.MoveNext
			$iCount2 += 1
		Until $oAD_RecordSet.EOF
	Else
		Local $aAD_Objects[$iCount1 + 1]
		$aAD_Objects[0] = UBound($aAD_Objects) - 1
		$iCount2 = 1
		$oAD_RecordSet.MoveFirst
		Do
			If IsArray($oAD_RecordSet.Fields($sAD_DataToRetrieve).Value) Then
				$aTemp = $oAD_RecordSet.Fields($sAD_DataToRetrieve).Value
				$aAD_Objects[$iCount2] = _ArrayToString($aTemp)
			Else
				$aAD_Objects[$iCount2] = $oAD_RecordSet.Fields($sAD_DataToRetrieve).Value
			EndIf
			$oAD_RecordSet.MoveNext
			$iCount2 += 1
		Until $oAD_RecordSet.EOF
	EndIf
	$__oAD_Command.Properties("Sort On") = "" ; Reset sort property
	Return $aAD_Objects

EndFunc   ;==>_AD_GetObjectsInOU

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetAllOUs
; Description ...: Retrieves an array of OUs. The paths are separated by the '\' character.
; Syntax.........: _AD_GetAllOUs([$sAD_Root = ""[, $sAD_Separator = "\"[, $iAD_Select = 0]]])
; Parameters ....: $sAD_Root      - Optional: OU (FQDN) where to start in the AD tree (default = "", equals "start at the AD root")
;                  $sAD_Separator - Optional: Single character to separate the OU names (default = "\")
;                  $iAD_Select    - Optional: Which objects should be returned in the result (default = 0)
;                  |0 - Return OUs (Organizational Units) (default)
;                  |1 - Return CNs (Containers)
;                  |2 - Return OUs + CNs
; Return values .: Success - One-based two dimensional array of OUs starting with the given OU. The paths are separated by "\"
;                  |0 - ... \name of grandfather OU\name of father OU\name of son OU
;                  |1 - Distinguished Name (FQDN) of the son OU
;                  Failure - "", sets @error to:
;                  |1 - No OUs found
;                  |2 - Specified $sAD_Root does not exist
;                  |3 - $iAD_Select is not an integer or < 0 or > 2
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: If an OU contains spaces the sorting is wrong and might lead to problems in further processing.
;                  Please have a look at http://www.autoitscript.com/forum/topic/106163-active-directory-udf/page__view__findpost__p__943892
; Related .......: _AD_GetObjectsInOU
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetAllOUs($sAD_Root = "", $sAD_Separator = "\", $iAD_Select = 0)

	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(2, 0, "")
	EndIf
	If Not IsInt($iAD_Select) Or $iAD_Select < 0 Or $iAD_Select > 2 Then Return SetError(3, 0, "")
	If $sAD_Separator <= " " Or StringLen($sAD_Separator) > 1 Then $sAD_Separator = "\"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;"
	Switch $iAD_Select
		Case  0
			$__oAD_Command.CommandText = $__oAD_Command.CommandText & "(objectCategory=organizationalUnit);distinguishedName;subtree"
		Case 1
			$__oAD_Command.CommandText = $__oAD_Command.CommandText & "(objectCategory=container);distinguishedName;subtree"
		Case 2
			$__oAD_Command.CommandText = $__oAD_Command.CommandText & "(|(objectCategory=organizationalUnit)(objectCategory=container));distinguishedName;subtree"
	EndSwitch
	Local $oAD_RecordSet = $__oAD_Command.Execute
	Local $iCount1 = $oAD_RecordSet.RecordCount
	If $iCount1 = 0 Then Return SetError(1, 0, "")
	Local $aAD_OUs[$iCount1 + 1][2]
	Local $iCount2 = 1, $aAD_TempOU
	$oAD_RecordSet.MoveFirst
	Do
		$aAD_OUs[$iCount2][1] = $oAD_RecordSet.Fields("distinguishedName").Value
		$aAD_OUs[$iCount2][0] = "," & StringTrimRight($aAD_OUs[$iCount2][1], StringLen($sAD_DNSDomain) + 1)
		$aAD_TempOU = StringSplit($aAD_OUs[$iCount2][0], "," & StringLeft($aAD_OUs[$iCount2][1], 3), 1) ; Split at ",OU=" or ",CN="
		_ArrayReverse($aAD_TempOU)
		$aAD_OUs[$iCount2][0] = StringTrimRight(_ArrayToString($aAD_TempOU, $sAD_Separator), 3)
		$iCount2 += 1
		$oAD_RecordSet.MoveNext
	Until $oAD_RecordSet.EOF
	_ArraySort($aAD_OUs)
	$aAD_OUs[0][0] = UBound($aAD_OUs, 1) - 1
	$aAD_OUs[0][1] = 2
	Return $aAD_OUs

EndFunc   ;==>_AD_GetAllOUs

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListDomainControllers
; Description ...: Enumerates all Domain Controllers (returns information about: Domain Controller, site, subnet and Global Catalog).
; Syntax.........: _AD_ListDomainControllers([$bAD_ListRO = False[, $bAD_ListGC = False]])
; Parameters ....: $bAD_ListRO - Optional: If set to True only returns RODC (read only domain controllers) (default = False)
;                  $bAD_ListGC - Optional: If set to True queries the DC for a Global Catalog. Disabled for performance reasons (default = False)
; Return values .: Success - One-based two dimensional array with the following information:
;                  |0 - Domain Controller: Name
;                  |1 - Domain Controller: Distinguished Name (FQDN)
;                  |2 - Domain Controller: DNS host name
;                  |3 - Site: Name
;                  |4 - Site: Distinguished Name (FQDN)
;                  |5 - Site: List of subnets that can connect to the site using this DC in the format x.x.x.x/mask - multiple subnets are separated by comma
;                  |6 - Global Catalog: Set to True if the DC is a Global Catalog (only if flag $bAD_ListGC = True. If False then "" is returned)
;                  Failure - "", sets @error to:
;                  |1 - No Domain Controllers found. @extended is set to the error returned by LDAP
; Author ........: water (based on VB functions by Richard L. Mueller)
; Modified.......:
; Remarks .......: This function only lists writeable DCs (default). To list RODC (read only DCs) use parameter $bAD_ListRO
; Related .......:
; Link ..........: http://www.rlmueller.net/Enumerate%20DCs.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListDomainControllers($bAD_ListRO = False, $bAD_ListGC = False)

	Local $oAD_DC, $oAD_Site, $oAD_Result
	Local Const $NTDSDSA_OPT_IS_GC = 1
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Configuration & ">;(objectClass=nTDSDSA);ADsPath;subtree"
	If $bAD_ListRO Then $__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Configuration & ">;(objectClass=nTDSDSARO);ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	; The parent object of each object with objectClass=nTDSDSA is a Domain
	; Controller. The parent of each Domain Controller is a "Servers"
	; container, and the parent of this container is the "Site" container.
	$oAD_RecordSet.MoveFirst
	Local $aAD_Result[1][7], $iCount1 = 1, $aAD_SubNet, $aAD_Temp, $sAD_Temp
	Do
		ReDim $aAD_Result[$iCount1 + 1][7]
		$oAD_Result = __AD_ObjGet($oAD_RecordSet.Fields("AdsPath").Value)
		$oAD_DC = __AD_ObjGet($oAD_Result.Parent)
		$aAD_Result[$iCount1][0] = $oAD_DC.Get("Name")
		$aAD_Result[$iCount1][1] = $oAD_DC.serverReference
		$aAD_Result[$iCount1][2] = $oAD_DC.DNSHostName
		$oAD_Result = __AD_ObjGet($oAD_DC.Parent)
		$oAD_Site = __AD_ObjGet($oAD_Result.Parent)
		$aAD_Result[$iCount1][3] = StringMid($oAD_Site.Name, 4)
		$aAD_Result[$iCount1][4] = $oAD_Site.distinguishedName
		$aAD_SubNet = $oAD_Site.GetEx("siteObjectBL")
		For $iCount2 = 0 To UBound($aAD_SubNet) - 1
			$aAD_Temp = StringSplit($aAD_SubNet[$iCount2], ",")
			$sAD_Temp = StringMid($aAD_Temp[1], 4)
			If $iCount2 = 0 Then
				$aAD_Result[$iCount1][5] = $sAD_Temp
			Else
				$aAD_Result[$iCount1][5] = $aAD_Result[$iCount1][5] & "," & $sAD_Temp
			EndIf
		Next
		If $bAD_ListGC Then
			; Is the DC a GC? Taken from: http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/computermanagement/ad/
			Local $oAD_DCRootDSE = __AD_ObjGet("LDAP://" & $oAD_DC.DNSHostName & "/rootDSE")
			Local $sAD_DsServiceDN = $oAD_DCRootDSE.Get("dsServiceName")
			Local $oAD_DsRoot = __AD_ObjGet("LDAP://" & $oAD_DC.DNSHostName & "/" & $sAD_DsServiceDN)
			Local $iAD_DCOptions = $oAD_DsRoot.Get("options")
			If BitAND($iAD_DCOptions, $NTDSDSA_OPT_IS_GC) = 1 Then
				$aAD_Result[$iCount1][6] = True
			Else
				$aAD_Result[$iCount1][6] = False
			EndIf
		EndIf
		$oAD_RecordSet.MoveNext
		$iCount1 += 1
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	$aAD_Result[0][0] = UBound($aAD_Result, 1) - 1
	$aAD_Result[0][1] = UBound($aAD_Result, 2)
	Return $aAD_Result

EndFunc   ;==>_AD_ListDomainControllers

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListRootDSEAttributes
; Description ...: Returns a one-based array of the RootDSE Atributes.
; Syntax.........: _AD_ListRootDSEAttributes()
; Parameters ....:
; Return values .: Success - Returns an one-based one dimensional array of the following RootDSE attributes. Multi-valued attributes are as multiple lines.
;                  |1 - configurationNamingContext: Specifies the distinguished name for the configuration container.
;                  |2 - currentTime: Specifies the current time set on this directory server in Coordinated Universal Time format.
;                  |3 - defaultNamingContext: Specifies the distinguished name of the domain that this directory server is a member.
;                  |4 - dnsHostName: Specifies the DNS address for this directory server.
;                  |5 - domainControllerFunctionality: Specifies the functional level of this domain controller. Values can be:
;                       0 - Windows 2000 Mode
;                       2 - Windows Server 2003 Mode
;                       3 - Windows Server 2008 Mode
;                       4 - Windows Server 2008 R2 Mode
;                  |6 - domainFunctionality: Specifies the functional level of the domain. Values can be:
;                       0 - Windows 2000 Domain Mode
;                       1 - Windows Server 2003 Interim Domain Mode
;                       2 - Windows Server 2003 Domain Mode
;                       3 - Windows Server 2008 Domain Mode
;                       4 - Windows Server 2008 R2 Domain Mode
;                  |7 - dsServiceName: Specifies the distinguished name of the NTDS settings object for this directory server.
;                  |8 - forestFunctionality: Specifies the functional level of the forest. Values can be:
;                       0 - Windows 2000 Forest Mode
;                       1 - Windows Server 2003 Interim Forest Mode
;                       2 - Windows Server 2003 Forest Mode
;                       3 - Windows Server 2008 Forest Mode
;                       4 - Windows Server 2008 R2 Forest Mode
;                  |9 - highestCommittedUSN: Specifies the highest update sequence number (USN) on this directory server. Used by directory replication.
;                  |10 - isGlobalCatalogReady: Specifies Global Catalog operational status. Values can be either "True" or "False".
;                  |11 - isSynchronized: Specifies directory server synchronisation status. Values can be either "True" or "False".
;                  |12 - LDAPServiceName: Specifies the Service Principal Name (SPN) for the LDAP server. Used for mutual authentication.
;                  |13 - namingContexts: A multi-valued attribute that specifies the distinguished names for all naming contexts stored on this directory server.
;                  +By default, a Windows 2000 domain controller has at least three naming contexts: Schema, Configuration, and the domain which the server is a member of.
;                  |14 - rootDomainNamingContext: Specifies the distinguished name for the first domain in the forest that this directory server is a member of.
;                  |15 - schemaNamingContext: Specifies the distinguished name for the schema container.
;                  |16 - serverName: Specifies the distinguished name of the server object for this directory server in the configuration container.
;                  |17 - subschemaSubentry: Specifies the distinguished name for the subSchema object. The subSchema object specifies properties that expose the supported attributes
;                  +(in the attributeTypes property) and classes (in the objectClasses property).
;                  |18 - supportedCapabilities: multi-valued attribute that specifies the capabilities supported by this directory server.
;                  |19 - supportedControl: A multi-valued attribute that specifies the extension control OIDs supported by this directory server.
;                  |20 - supportedLDAPPolicies: A multi-valued attribute that specifies the names of the supported LDAP management policies.
;                  |21 - supportedLDAPVersion: A multi-valued attribute that specifies the LDAP versions (specified by major version number) supported by this directory server.
;                  |22 - supportedSASLMechanisms: Specifies the security mechanisms supported for SASL negotiation (see LDAP RFCs). By default, GSSAPI is supported.
; Author ........: water
; Modified.......:
; Remarks .......: In LDAP 3.0, rootDSE is defined as the root of the directory data tree on a directory server.
;                  The rootDSE is not part of any namespace. The purpose of the rootDSE is to provide data about the directory server.
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/cc223254(v=PROT.13).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListRootDSEAttributes()

	Return _AD_GetObjectProperties($__oAD_RootDSE)

EndFunc   ;==>_AD_ListRootDSEAttributes

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListRoleOwners
; Description ...: Returns a one-based array of FSMO (Flexible Single Master Operation) Role Owners.
; Syntax.........: _AD_ListRoleOwners()
; Parameters ....:
; Return values .: Success - Returns an one-based one dimensional array of FSMO Role Owners. The array contains:
;                  |1 - Domain PDC FSMO
;                  |2 - Domain Rid FSMO
;                  |3 - Domain Infrastructure FSMO
;                  |4 - Forest-wide Schema FSMO
;                  |5 - Forest-wide Domain naming FSMO
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: http://www.tools4net.de/doc/ad2.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListRoleOwners()

	Local $aAD_Roles[6]

	; PDC FSMO
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectClass=domainDNS)(fSMORoleOwner=*));adsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	Local $oAD_FSM = ObjGet($oAD_RecordSet.fields(0).value)
	Local $oAD_CompNTDS = ObjGet("LDAP://" & $sAD_HostServer & "/" & $oAD_FSM.FSMORoleOwner)
	Local $oAD_Comp = ObjGet($oAD_CompNTDS.Parent)
	$aAD_Roles[1] = $oAD_Comp.dnsHostname

	; Rid FSMO
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectClass=rIDManager) (fSMORoleOwner=*));adsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute
	$oAD_FSM = ObjGet($oAD_RecordSet.fields(0).value)
	$oAD_CompNTDS = ObjGet("LDAP://" & $sAD_HostServer & "/" & $oAD_FSM.FSMORoleOwner)
	$oAD_Comp = ObjGet($oAD_CompNTDS.Parent)
	$aAD_Roles[2] = $oAD_Comp.dnsHostname

	; Infrastructure FSMO
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectClass=infrastructureUpdate) (fSMORoleOwner=*));adsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute
	$oAD_FSM = ObjGet($oAD_RecordSet.fields(0).value)
	$oAD_CompNTDS = ObjGet("LDAP://" & $sAD_HostServer & "/" & $oAD_FSM.FSMORoleOwner)
	$oAD_Comp = ObjGet($oAD_CompNTDS.Parent)
	$aAD_Roles[3] = $oAD_Comp.dnsHostname

	; Schema FSMO
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $__oAD_RootDSE.Get("schemaNamingContext") & ">;(&(objectClass=dMD) (fSMORoleOwner=*));adsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute
	$oAD_FSM = ObjGet($oAD_RecordSet.fields(0).value)
	$oAD_CompNTDS = ObjGet("LDAP://" & $sAD_HostServer & "/" & $oAD_FSM.FSMORoleOwner)
	$oAD_Comp = ObjGet($oAD_CompNTDS.Parent)
	$aAD_Roles[4] = $oAD_Comp.dnsHostname

	; Domain Naming FSMO
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $__oAD_RootDSE.Get("configurationNamingContext") & ">;(&(objectClass=crossRefContainer)(fSMORoleOwner=*));adsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute
	$oAD_FSM = ObjGet($oAD_RecordSet.fields(0).value)
	$oAD_CompNTDS = ObjGet("LDAP://" & $sAD_HostServer & "/" & $oAD_FSM.FSMORoleOwner)
	$oAD_Comp = ObjGet($oAD_CompNTDS.Parent)
	$aAD_Roles[5] = $oAD_Comp.dnsHostname

	$aAD_Roles[0] = 5
	Return $aAD_Roles

EndFunc   ;==>_AD_ListRoleOwners

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetLastLoginDate
; Description ...: Returns the lastlogin information from all DCs using the SamAccountName.
; Syntax.........: _AD_GetLastLoginDate([$sAD_User = @Username[, $sAD_Site = ""[, $aAD_DCList = ""]]])
; Parameters ....: $sAD_User   - Optional: SamAccountName of a user account to get the last login date (default = @Username).
;                  $sAD_Site   - Optional: Only query DCs that belong to this site(s) (default = all sites).
;                  +This can be a single site or a list of sites separated by commas
;                  $aAD_DCList - Optional: one-based two dimensional array of Domain Controllers as returned by function _AD_ListDomainControllers (default = "")
; Return values .: Success - Last login date returned as YYYYMMDDHHMMSS. @extended is set to the total number of Domain Controllers.
;                  +@error could be > 0 and contains the number of DCs that could not be reached or returns no data
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User could not be found. @extended = 0
;                  |2 - $sAD_User has never logged in to the domain. @extended = 0
;                  |3 - $aAD_DCList has to be an array or blank
;                  |4 - $aAD_DCList has to be a 2-dimensional array
;                  Warning - Last login date returned as YYYYMMDDHHMMSS (see Success), sets @error and @extended to:
;                  |x - Number of DCs which could not be reached. Result is returned from all available DCs. @extended is set to the total number of Domain Controllers
; Author ........: Jonathan Clelland
; Modified.......: water, Stephane
; Remarks .......: If it takes (too) long to get a result either some DCs are down or you have too many DCs in your AD.
;                  +Case one: Please check @error and @extended as described above
;                  +Case two: Specify parameter $sAD_Site to reduce the number of DCs to query and/or retrieve the list of DCs yourself and pass the array as parameter 3
; Related .......:
; Link ..........: http://blogs.technet.com/b/askds/archive/2009/04/15/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetLastLoginDate($sAD_User = @UserName, $sAD_Site = "", $aAD_DCList = "")

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If Not IsArray($aAD_DCList) And $aAD_DCList <> "" Then Return SetError(3, 0, 0)
	If IsArray($aAD_DCList) And UBound($aAD_DCList, 0) <> 2 Then Return SetError(4, 0, 0)
	If $aAD_DCList = "" Then $aAD_DCList = _AD_ListDomainControllers()
	Local $aAD_Site, $sAD_SingleDC, $bAD_WasIn
	; Delete all DCs not belonging to the specified site
	$aAD_Site = StringSplit($sAD_Site, ",", 2)
	If UBound($aAD_Site) > 0 And $aAD_Site[0] <> "" Then
		For $iAD_Count1 = $aAD_DCList[0][0] To 1 Step -1
			$bAD_WasIn = False
			For $sAD_SingleDC In $aAD_Site
				If $aAD_DCList[$iAD_Count1][3] = $sAD_SingleDC Then $bAD_WasIn = True
			Next
			If Not $bAD_WasIn Then _ArrayDelete($aAD_DCList, $iAD_Count1)
		Next
		$aAD_DCList[0][0] = UBound($aAD_DCList, 1) - 1
	EndIf
	; Get LastLogin from all DCs
	Local $aAD_Result[$aAD_DCList[0][0] + 1]
	Local $sAD_LDAPEntry, $oAD_Object, $oAD_RecordSet
	Local $iAD_Error1 = 0, $iAD_Error2 = 0
	For $iCount1 = 1 To $aAD_DCList[0][0]
		If Ping($aAD_DCList[$iCount1][2]) = 0 Then
			$iAD_Error1 += 1
			ContinueLoop
		EndIf
		$__oAD_Command.CommandText = "<LDAP://" & $aAD_DCList[$iCount1][2] & "/" & $sAD_DNSDomain & ">;(sAMAccountName=" & $sAD_User & ");ADsPath;subtree"
		$oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
		; -2147352567 or 0x80020009 is returned when the service is not operational
		If @error = -2147352567 Or $oAD_RecordSet.RecordCount = 0 Then
			$iAD_Error1 += 1
		Else
			$sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
			$oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
			$aAD_Result[$iCount1] = $oAD_Object.LastLogin
			; -2147352567 or 0x80020009 is returned when the attribute "LastLogin" isn't defined on this DC
			If @error = -2147352567 Then $iAD_Error2 += 1
			$oAD_Object.PurgePropertyList
		EndIf
	Next
	_ArraySort($aAD_Result, 1, 1)
	; If error count equals the number of DCs then the user has never logged in
	If $iAD_Error2 = $aAD_DCList[0][0] Then Return SetError(2, 0, 0)
	Return SetError($iAD_Error1, $aAD_DCList[0][0], $aAD_Result[1])

EndFunc   ;==>_AD_GetLastLoginDate

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_IsObjectDisabled
; Description ...: Returns 1 if the object (user account, computer account) is disabled.
; Syntax.........: _AD_IsObjectDisabled([$sAD_Object = @Username])
; Parameters ....: $sAD_Object - Optional: Object to check (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - 1, Specified object is disabled
;                  Failure - 0, sets @error to:
;                  |0 - $sAD_Object is not disabled
;                  |1 - $sAD_Object could not be found
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: A $ sign must be appended to the computer name to create a correct sAMAccountName e.g. @ComputerName & "$"
; Related .......: _AD_DisableObject, _AD_EnableObject, _AD_GetObjectsDisabled
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_IsObjectDisabled($sAD_Object = @UserName)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, 0)
	Local $iAD_UAC = _AD_GetObjectAttribute($sAD_Object, "userAccountControl")
	If BitAND($iAD_UAC, $ADS_UF_ACCOUNTDISABLE) = $ADS_UF_ACCOUNTDISABLE Then Return 1
	Return 0

EndFunc   ;==>_AD_IsObjectDisabled

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_IsObjectLocked
; Description ...: Returns 1 if the object (user account, computer account) is locked.
; Syntax.........: _AD_IsObjectLocked([$sAD_Object = @Username])
; Parameters ....: $sAD_Object - Optional: Object to check (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - 1, Specified object is locked, sets @error to:
;                  |x - number of minutes till the account is unlocked. -1 means the account has to be unlocked manually by an admin
;                  Failure - 0, sets @error to:
;                  |0 - $sAD_Object is not locked
;                  |1 - $sAD_Object could not be found
; Author ........: water
; Modified.......:
; Remarks .......: A $ sign must be appended to the computer name to create a correct sAMAccountName e.g. @ComputerName & "$"
;                  LockoutTime contains the timestamp when the object was locked. This value is not reset until the user/computer logs on again.
;                  LockoutTime could be > 0 even when the lockout already has expired.
; Related .......: _AD_GetObjectsLocked, _AD_UnlockObject
; Link ..........: http://www.pcreview.co.uk/forums/thread-1350048.php, http://www.rlmueller.net/IsUserLocked.htm, http://technet.microsoft.com/en-us/library/cc780271%28WS.10%29.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_IsObjectLocked($sAD_Object = @UserName)

	;-------------------------------------------------------------
	; HINT: To enhance performance this can also be written as:
	; 	$oUser = ObjGet("WinNT://<Domain>/<User>")
	;	ConsoleWrite("Locked: " & $oUser.IsAccountLocked & @CRLF)
	;-------------------------------------------------------------
	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	Local $oAD_LockoutTime = $oAD_Object.LockoutTime
	; Object is not locked out
	If Not IsObj($oAD_LockoutTime) Then Return
	; Calculate lockout time (UTC)
	Local $sAD_LockoutTime = _DateAdd("s", Int(__AD_LargeInt2Double($oAD_LockoutTime.LowPart, $oAD_LockoutTime.HighPart) / (10000000)), "1601/01/01 00:00:00")
	; Object is not locked out
	If $sAD_LockoutTime = "1601/01/01 00:00:00" Then Return
	; Get password info - Account Lockout Duration
	Local $aAD_Temp = _AD_GetPasswordInfo($sAD_Object)
	; if lockout duration is 0 (= unlock manually by admin needed) then no calculation is necessary. Set @error to -1 (minutes till the account is unlocked)
	If $aAD_Temp[5] = 0 Then Return SetError(-1, 0, 1)
	; Calculate when the lockout will be reset
	Local $sAD_ResetLockoutTime = _DateAdd("n", $aAD_Temp[5], $sAD_LockoutTime)
	; Compare to current date/time (UTC)
	Local $sAD_Now = _Date_Time_GetSystemTime()
	$sAD_Now = _Date_Time_SystemTimeToDateTimeStr($sAD_Now, 1)
	If $sAD_ResetLockoutTime >= $sAD_Now Then Return SetError(_DateDiff("n", $sAD_Now, $sAD_ResetLockoutTime), 0, 1)
	Return

EndFunc   ;==>_AD_IsObjectLocked

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_IsPasswordExpired
; Description ...: Returns 1 if the password of the user or computer account has expired.
; Syntax.........: _AD_IsPasswordExpired([$sAD_Account = @Username])
; Parameters ....: $sAD_Account - Optional: User or computer account to check (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - 1, The password of the specified account has expired
;                  Failure - 0, sets @error to:
;                  |0 - Password for $sAD_Account has not expired
;                  |1 - $sAD_Account could not be found
;                  |x - Error as returned by function _AD_GetPasswordInfo
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_IsPasswordExpired($sAD_Account = @UserName)

	If Not _AD_ObjectExists($sAD_Account) Then Return SetError(1, 0, 0)
	Local $aAD_Temp = _AD_GetPasswordInfo($sAD_Account)
	If @error <> 0 Then SetError(@error, 0, 0)
	If $aAD_Temp[11] <= _NowCalc() Then Return 1
	Return

EndFunc   ;==>_AD_IsPasswordExpired

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectsDisabled
; Description ...: Returns an array with FQDNs of disabled objects (user accounts, computer accounts).
; Syntax.........: _AD_GetObjectsDisabled([$sAD_Class = "user"[, $sAD_Root = ""]])
; Parameters ....: $sAD_Class - Optional: Specifies if disabled user accounts or computer accounts should be returned (default = "user").
;                  |"user"     - Returns objects of category "user"
;                  |"computer" - Returns objects of category "computer"
;                  $sAD_Root  - Optional: FQDN of the OU where the search should start (default = "" = search the whole tree)
; Return values .: Success - array of user or computer account FQDNs
;                  Failure - "", sets @error to:
;                  |1 - $sAD_Class is invalid. Values can be "computer" or "user"
;                  |2 - Specified $sAD_Root does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsObjectDisabled, _AD_DisableObject, _AD_EnableObject
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectsDisabled($sAD_Class = "user", $sAD_Root = "")

	If $sAD_Class <> "user" And $sAD_Class <> "computer" Then Return SetError(1, 0, "")
	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(2, 0, "")
	EndIf
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;(&(objectcategory=" & $sAD_Class & ")(userAccountControl:1.2.840.113556.1.4.803:=" & _
			$ADS_UF_ACCOUNTDISABLE & "));distinguishedName,objectcategory;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	Local $aAD_FQDN[$oAD_RecordSet.RecordCount + 1]
	$aAD_FQDN[0] = $oAD_RecordSet.RecordCount
	Local $iCount1 = 1
	While Not $oAD_RecordSet.EOF
		$aAD_FQDN[$iCount1] = $oAD_RecordSet.Fields(0).Value
		$iCount1 += 1
		$oAD_RecordSet.MoveNext
	WEnd
	Return $aAD_FQDN

EndFunc   ;==>_AD_GetObjectsDisabled

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectsLocked
; Description ...: Returns an array of FQDNs of locked (user and/or, computer accounts), lockout time and minutes remaining in locked state.
; Syntax.........: _AD_GetObjectsLocked([$sAD_Class = "user"[, $sAD_Root = ""]])
; Parameters ....: $sAD_Class - Optional: Specifies if locked user accounts or computer accounts should be returned (default = "user").
;                  |"user" - Returns objects of category "user"
;                  |"computer" - Returns objects of category "computer"
;                  $sAD_Root  - Optional: FQDN of the OU where the search should start (default = "" = search the whole tree)
; Return values .: Success - Returns a one-based two dimensional array with the following information:
;                  |0 - FQDN of the locked object
;                  |1 - lockout time YYYY/MM/DD HH:MM:SS in local time of the calling user
;                  |2 - Minutes until the object will be unlocked
;                  Failure - "", sets @error to:
;                  |1 - $sAD_Class is invalid. Should be "computer" or "user"
;                  |2 - No locked objects found
;                  |3 - Specified $sAD_Root does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: LockoutTime contains the timestamp when the object was locked. This value is not reset until the user/computer logs on again.
;                  LockoutTime could be > 0 even when the lockout has already expired.
; Related .......: _AD_IsObjectLocked, _AD_UnlockObject
; Link ..........: http://technet.microsoft.com/en-us/library/cc780271%28WS.10%29.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectsLocked($sAD_Class = "user", $sAD_Root = "")

	If $sAD_Class <> "user" And $sAD_Class <> "computer" Then Return SetError(1, 0, "")
	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(3, 0, "")
	EndIf
	; Get all objects with lockouttime>=1
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;(&(objectcategory=" & $sAD_Class & ")(lockouttime>=1));distinguishedName;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(2, @error, "")
	Local $aAD_FQDN[$oAD_RecordSet.RecordCount + 1][3] = [[$oAD_RecordSet.RecordCount, 3]]
	Local $iCount1 = 1
	Local $aAD_Result
	While Not $oAD_RecordSet.EOF
		$aAD_FQDN[$iCount1][0] = $oAD_RecordSet.Fields(0).Value
		$iCount1 += 1
		$oAD_RecordSet.MoveNext
	WEnd
	; check if lockouttime has expired. If yes, delete from table
	For $iCount1 = $aAD_FQDN[0][0] To 1 Step -1
		If Not _AD_IsObjectLocked($aAD_FQDN[$iCount1][0]) Then
			_ArrayDelete($aAD_FQDN, $iCount1)
		Else
			$aAD_FQDN[$iCount1][2] = @error
			$aAD_Result = _AD_GetObjectProperties($aAD_FQDN[$iCount1][0], "lockouttime")
			$aAD_FQDN[$iCount1][1] = $aAD_Result[1][1]
		EndIf
	Next
	$aAD_FQDN[0][0] = UBound($aAD_FQDN) - 1
	If $aAD_FQDN[0][0] = 0 Then Return SetError(2, 0, "")
	Return $aAD_FQDN

EndFunc   ;==>_AD_GetObjectsLocked

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetPasswordExpired
; Description ...: Returns an array of FQDNs of user or computer accounts with expired passwords.
; Syntax.........: _AD_GetPasswordExpired([$sAD_Root = ""[, $bAD_NeverChanged = False]])
; Parameters ....: $sAD_Root         - Optional: FQDN of the OU where the search should start (default = "" = search the whole tree)
;                  $bAD_NeverChanged - Optional: If set to True returns all accounts who have never changed their password as well (default = False)
;                  $iAD_PasswordAge  - Optional: Takes the max. password age from the AD or uses this value if > 0
;                  $bAD_Computer     - Optional: If True queries computer accounts, if False queries user accounts (default = False)
; Return values .: Success - One-based two dimensional array of FQDNs of accounts with expired passwords
;                  |0 - FQDNs of accounts with expired password
;                  |1 - password last set YYYY/MM/DD HH:NMM:SS UTC
;                  |2 - password last set YYYY/MM/DD HH:NMM:SS local time of calling user
;                  Failure - "", sets @error to:
;                  |1 - No expired passwords found. @extended is set to the error returned by LDAP
;                  |2 - Specified $sAD_Root does not exist
;                  |3 - $iAD_PasswordAge is not numeric
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetPasswordExpired($sAD_Root = "", $bAD_NeverChanged = False, $iAD_PasswordAge = 0, $bAD_Computer = False)

	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(2, 0, "")
	EndIf
	If $iAD_PasswordAge <> 0 And Not IsNumber($iAD_PasswordAge) Then Return SetError(3, 0, "")
	Local $aAD_Temp = _AD_GetPasswordInfo()
	Local $sAD_DTExpire = _Date_Time_GetSystemTime() ; Get current date/time
	$sAD_DTExpire = _Date_Time_SystemTimeToDateTimeStr($sAD_DTExpire, 1) ; convert to system time
	If $iAD_PasswordAge <> 0 Then
		$sAD_DTExpire = _DateAdd("D", $iAD_PasswordAge * - 1, $sAD_DTExpire) ; substract maximum password age
	Else
		$sAD_DTExpire = _DateAdd("D", $aAD_Temp[1] * - 1, $sAD_DTExpire) ; substract maximum password age
	EndIf
	Local $iAD_DTExpire = _DateDiff("s", "1601/01/01 00:00:00", $sAD_DTExpire) * 10000000 ; convert to Integer8
	Local $sAD_DTStruct = DllStructCreate("dword low;dword high")
	Local $sAD_Temp, $iAD_Temp, $iAD_LowerDate = 110133216000000001 ; 110133216000000001 = 01/01/1959 00:00:00 UTC
	If $bAD_NeverChanged = True Then $iAD_LowerDate = 0
	Local $sAD_Category = "(objectCategory=Person)(objectClass=User)"
	If $bAD_Computer = True Then $sAD_Category = "(objectCategory=computer)"
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;(&" & $sAD_Category & _
			"(pwdLastSet<=" & Int($iAD_DTExpire) & ")(pwdLastSet>=" & $iAD_LowerDate & "));distinguishedName,pwdlastset,useraccountcontrol;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	Local $aAD_FQDN[$oAD_RecordSet.RecordCount + 1][3]
	$aAD_FQDN[0][0] = $oAD_RecordSet.RecordCount
	Local $iAD_Count = 1
	While Not $oAD_RecordSet.EOF
		$aAD_FQDN[$iAD_Count][0] = $oAD_RecordSet.Fields(0).Value
		$iAD_Temp = $oAD_RecordSet.Fields(1).Value
		If BitAND($oAD_RecordSet.Fields(2).Value, $ADS_UF_DONT_EXPIRE_PASSWD) <> $ADS_UF_DONT_EXPIRE_PASSWD Then
			DllStructSetData($sAD_DTStruct, "Low", $iAD_Temp.LowPart)
			DllStructSetData($sAD_DTStruct, "High", $iAD_Temp.HighPart)
			$sAD_Temp = _Date_Time_FileTimeToSystemTime(DllStructGetPtr($sAD_DTStruct))
			$aAD_FQDN[$iAD_Count][1] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp, 1)
			$sAD_Temp = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sAD_Temp))
			$aAD_FQDN[$iAD_Count][2] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp, 1)
		EndIf
		$iAD_Count += 1
		$oAD_RecordSet.MoveNext
	WEnd
	; Delete records with UAC set to password not expire
	$aAD_FQDN[0][0] = UBound($aAD_FQDN) - 1
	For $iAD_Count = $aAD_FQDN[0][0] To 1 Step -1
		If $aAD_FQDN[$iAD_Count][1] = "" Then _ArrayDelete($aAD_FQDN, $iAD_Count)
	Next
	$aAD_FQDN[0][0] = UBound($aAD_FQDN) - 1
	Return $aAD_FQDN

EndFunc   ;==>_AD_GetPasswordExpired

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetPasswordDontExpire
; Description ...: Returns an array of user account FQDNs where the password does not expire.
; Syntax.........: _AD_GetPasswordDontExpire([$sAD_Root = ""])
; Parameters ....: $sAD_Root - Optional: FQDN of the OU where the search should start (default = "" = search the whole tree)
; Return values .: Success - Array with FQDNs of user accounts for which the password does not expire
;                  Failure - "", sets @error to:
;                  |1 - No user accounts for which the password does not expire. @extended is set to the error returned by LDAP
;                  |2 - Specified $sAD_Root does not exist
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetPasswordDontExpire($sAD_Root = "")

	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(2, 0, "")
	EndIf
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;(&(objectcategory=user)(userAccountControl:1.2.840.113556.1.4.803:=" & _
			$ADS_UF_DONT_EXPIRE_PASSWD & "));distinguishedName;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	Local $aAD_FQDN[$oAD_RecordSet.RecordCount + 1]
	$aAD_FQDN[0] = $oAD_RecordSet.RecordCount
	Local $iCount1 = 1
	While Not $oAD_RecordSet.EOF
		$aAD_FQDN[$iCount1] = $oAD_RecordSet.Fields(0).Value
		$iCount1 += 1
		$oAD_RecordSet.MoveNext
	WEnd
	Return $aAD_FQDN

EndFunc   ;==>_AD_GetPasswordDontExpire

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetObjectProperties
; Description ...: Returns a two-dimensional array of all or selected properties and their values of an object in readable form.
; Syntax.........: _AD_GetObjectProperties([$sAD_Object = @UserName[, $sAD_Properties = ""]])
; Parameters ....: $sAD_Object - Optional: SamAccountName or FQDN of the object properties from (e.g. computer, user, group ...) (default = @Username)
;                  |Can be of type object as well. Useful to get properties for a schema or configuration object (see _AD_ListRootDSEAttributes)
;                  $sAD_Properties - Optional: Comma separated list of properties to return (default = "" = return all properties)
; Return values .: Success - Returns a two-dimensional array with all properties and their values of an object in readable form
;                  Failure - "" or property name, sets @error to:
;                  |1 - $sAD_Object could not be found
;                  |2 - No values for the specified property. The property in error is returned as the function result
; Author ........: Sundance
; Modified.......: water
; Remarks .......: Dates are returned in format: YYYY/MM/DD HH:MM:SS local time of the calling user (AD stores all dates in UTC - Universal Time Coordinated)
;                  Exception: AD internal dates like "whenCreated", "whenChanged" and "dSCorePropagationData". They are returned as UTC
;                  NT Security Descriptors are returned as: Control:nn, Group:Domain\Group, Owner:Domain\Group, Revision:nn
;                  No error is returned if there are properties in $sAD_Properties that are not available for the selected object
;+
;                  Properties are returned in alphabetical order. If $sAD_Properties is set to "samaccountname,displayname" the returned array will contain
;                  displayname as the first and samaccountname as the second row.
; Related .......:
; Link ..........: http://www.autoitscript.com/forum/index.php?showtopic=49627&view=findpost&p=422402, http://msdn.microsoft.com/en-us/library/ms675090(VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetObjectProperties($sAD_Object = @UserName, $sAD_Properties = "")

	Local $aAD_ObjectProperties[1][2], $oAD_Object
	Local $oAD_Item, $oAD_PropertyEntry, $oAD_Value, $iCount3, $xAD_Dummy
	; Data Type Mapping between Active Directory and LDAP
	; http://msdn.microsoft.com/en-us/library/aa772375(VS.85).aspx
	Local Const $ADSTYPE_DN_STRING = 1
	Local Const $ADSTYPE_CASE_IGNORE_STRING = 3
	Local Const $ADSTYPE_BOOLEAN = 6
	Local Const $ADSTYPE_INTEGER = 7
	Local Const $ADSTYPE_OCTET_STRING = 8
	Local Const $ADSTYPE_UTC_TIME = 9
	Local Const $ADSTYPE_LARGE_INTEGER = 10
	Local Const $ADSTYPE_NT_SECURITY_DESCRIPTOR = 25
	Local Const $ADSTYPE_UNKNOWN = 26
	Local $aAD_SAMAccountType[12][2] = [["DOMAIN_OBJECT", 0x0],["GROUP_OBJECT", 0x10000000],["NON_SECURITY_GROUP_OBJECT", 0x10000001], _
			["ALIAS_OBJECT", 0x20000000],["NON_SECURITY_ALIAS_OBJECT", 0x20000001],["USER_OBJECT", 0x30000000],["NORMAL_USER_ACCOUNT", 0x30000000], _
			["MACHINE_ACCOUNT", 0x30000001],["TRUST_ACCOUNT", 0x30000002],["APP_BASIC_GROUP", 0x40000000],["APP_QUERY_GROUP", 0x40000001], _
			["ACCOUNT_TYPE_MAX", 0x7fffffff]]
	Local $aAD_UAC[21][2] = [[0x00000001, "SCRIPT"],[0x00000002, "ACCOUNTDISABLE"],[0x00000008, "HOMEDIR_REQUIRED"],[0x00000010, "LOCKOUT"],[0x00000020, "PASSWD_NOTREQD"], _
			[0x00000040, "PASSWD_CANT_CHANGE"],[0x00000080, "ENCRYPTED_TEXT_PASSWORD_ALLOWED"],[0x00000100, "TEMP_DUPLICATE_ACCOUNT"],[0x00000200, "NORMAL_ACCOUNT"], _
			[0x00000800, "INTERDOMAIN_TRUST_ACCOUNT"],[0x00001000, "WORKSTATION_TRUST_ACCOUNT"],[0x00002000, "SERVER_TRUST_ACCOUNT"],[0x00010000, "DONT_EXPIRE_PASSWD"], _
			[0x00020000, "MNS_LOGON_ACCOUNT"],[0x00040000, "SMARTCARD_REQUIRED"],[0x00080000, "TRUSTED_FOR_DELEGATION"],[0x00100000, "NOT_DELEGATED"], _
			[0x00200000, "USE_DES_KEY_ONLY"],[0x00400000, "DONT_REQUIRE_PREAUTH"],[0x00800000, "PASSWORD_EXPIRED"],[0x01000000, "TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION"]]

	$sAD_Properties = "," & StringReplace($sAD_Properties, " ", "") & ","
	If Not IsObj($sAD_Object) Then
		If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, "")
		Local $sAD_Property = "sAMAccountName"
		If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
		$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
		Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
		Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
		$oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object
	Else
		$oAD_Object = $sAD_Object
	EndIf
	$oAD_Object.GetInfo()
	Local $iCount1 = $oAD_Object.PropertyCount()
	For $iCount2 = 0 To $iCount1 - 1
		$oAD_Item = $oAD_Object.Item($iCount2)
		If Not ($sAD_Properties = ",," Or StringInStr($sAD_Properties, "," & $oAD_Item.Name & ",") > 0) Then ContinueLoop
		$oAD_PropertyEntry = $oAD_Object.GetPropertyItem($oAD_Item.Name, $ADSTYPE_UNKNOWN)
		If Not IsObj($oAD_PropertyEntry) Then
			Return SetError(2, 0, $oAD_Item.Name)
		Else
			For $vAD_PropertyValue In $oAD_PropertyEntry.Values
				ReDim $aAD_ObjectProperties[UBound($aAD_ObjectProperties, 1) + 1][2]
				$iCount3 = UBound($aAD_ObjectProperties, 1) - 1
				$aAD_ObjectProperties[$iCount3][0] = $oAD_Item.Name
				If $oAD_Item.ADsType = $ADSTYPE_CASE_IGNORE_STRING Then
					$aAD_ObjectProperties[$iCount3][1] = $vAD_PropertyValue.CaseIgnoreString
				ElseIf $oAD_Item.ADsType = $ADSTYPE_INTEGER Then
					If $oAD_Item.Name = "sAMAccountType" Then
						For $iCount4 = 0 To 11
							If $vAD_PropertyValue.Integer = $aAD_SAMAccountType[$iCount4][1] Then
								$aAD_ObjectProperties[$iCount3][1] = $aAD_SAMAccountType[$iCount4][0]
								ExitLoop
							EndIf
						Next
					ElseIf $oAD_Item.Name = "userAccountControl" Then
						$aAD_ObjectProperties[$iCount3][1] = $vAD_PropertyValue.Integer & " = "
						For $iCount4 = 0 To 20
							If BitAND($vAD_PropertyValue.Integer, $aAD_UAC[$iCount4][0]) = $aAD_UAC[$iCount4][0] Then
								$aAD_ObjectProperties[$iCount3][1] &= $aAD_UAC[$iCount4][1] & " - "
							EndIf
						Next
						If StringRight($aAD_ObjectProperties[$iCount3][1], 3) = " - " Then $aAD_ObjectProperties[$iCount3][1] = StringTrimRight($aAD_ObjectProperties[$iCount3][1], 3)
					Else
						$aAD_ObjectProperties[$iCount3][1] = $vAD_PropertyValue.Integer
					EndIf
				ElseIf $oAD_Item.ADsType = $ADSTYPE_LARGE_INTEGER Then
					If $oAD_Item.Name = "pwdLastSet" Or $oAD_Item.Name = "accountExpires" Or $oAD_Item.Name = "lastLogonTimestamp" Or $oAD_Item.Name = "badPasswordTime" Or $oAD_Item.Name = "lastLogon" Or $oAD_Item.Name = "lockoutTime" Then
						If $vAD_PropertyValue.LargeInteger.LowPart = 0 And $vAD_PropertyValue.LargeInteger.HighPart = 0 Then
							$aAD_ObjectProperties[$iCount3][1] = "1601/01/01 00:00:00"
						Else
							Local $sAD_Temp = DllStructCreate("dword low;dword high")
							DllStructSetData($sAD_Temp, "Low", $vAD_PropertyValue.LargeInteger.LowPart)
							DllStructSetData($sAD_Temp, "High", $vAD_PropertyValue.LargeInteger.HighPart)
							Local $sAD_Temp2 = _Date_Time_FileTimeToSystemTime(DllStructGetPtr($sAD_Temp))
							Local $sAD_Temp3 = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sAD_Temp2))
							$aAD_ObjectProperties[$iCount3][1] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp3, 1)
						EndIf
					Else
						$aAD_ObjectProperties[$iCount3][1] = __AD_LargeInt2Double($vAD_PropertyValue.LargeInteger.LowPart, $vAD_PropertyValue.LargeInteger.HighPart)
					EndIf
				ElseIf $oAD_Item.ADsType = $ADSTYPE_OCTET_STRING Then
					$xAD_Dummy = DllStructCreate("byte[56]")
					DllStructSetData($xAD_Dummy, 1, $vAD_PropertyValue.OctetString)
					; objectSID etc. See: http://msdn.microsoft.com/en-us/library/aa379597(VS.85).aspx
					; objectGUID etc. See: http://www.autoitscript.com/forum/index.php?showtopic=106163&view=findpost&p=767558
					If _Security__IsValidSid(DllStructGetPtr($xAD_Dummy)) Then
						$aAD_ObjectProperties[$iCount3][1] = _Security__SidToStringSid(DllStructGetPtr($xAD_Dummy)) ; SID
					Else
						$aAD_ObjectProperties[$iCount3][1] = _WinAPI_StringFromGUID(DllStructGetPtr($xAD_Dummy)) ; GUID
					EndIf
				ElseIf $oAD_Item.ADsType = $ADSTYPE_DN_STRING Then
					$aAD_ObjectProperties[$iCount3][1] = $vAD_PropertyValue.DNString
				ElseIf $oAD_Item.ADsType = $ADSTYPE_UTC_TIME Then
					Local $iAD_DateTime = $vAD_PropertyValue.UTCTime
					$aAD_ObjectProperties[$iCount3][1] = StringLeft($iAD_DateTime, 4) & "/" & StringMid($iAD_DateTime, 5, 2) & "/" & StringMid($iAD_DateTime, 7, 2) & _
							" " & StringMid($iAD_DateTime, 9, 2) & ":" & StringMid($iAD_DateTime, 11, 2) & ":" & StringMid($iAD_DateTime, 13, 2)
				ElseIf $oAD_Item.ADsType = $ADSTYPE_BOOLEAN Then
					If $vAD_PropertyValue.Boolean = 0 Then
						$aAD_ObjectProperties[$iCount3][1] = "False"
					Else
						$aAD_ObjectProperties[$iCount3][1] = "True"
					EndIf
				ElseIf $oAD_Item.ADsType = $ADSTYPE_NT_SECURITY_DESCRIPTOR Then
					$oAD_Value = $vAD_PropertyValue.SecurityDescriptor
					$aAD_ObjectProperties[$iCount3][1] = "Control:" & $oAD_Value.Control & ", " & _
							"Group:" & $oAD_Value.Group & ", " & _
							"Owner:" & $oAD_Value.Owner & ", " & _
							"Revision:" & $oAD_Value.Revision
				Else
					$aAD_ObjectProperties[$iCount3][1] = "Has the unknown ADsType: " & $oAD_Item.ADsType
				EndIf
			Next
		EndIf
	Next
	$aAD_ObjectProperties[0][0] = UBound($aAD_ObjectProperties, 1) - 1
	$aAD_ObjectProperties[0][1] = UBound($aAD_ObjectProperties, 2)
	_ArraySort($aAD_ObjectProperties, 0, 1)
	Return $aAD_ObjectProperties

EndFunc   ;==>_AD_GetObjectProperties

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_CreateOU
; Description ...: Creates a child OU in the specified parent OU.
; Syntax.........: _AD_CreateOU($sAD_ParentOU, $sAD_OU)
; Parameters ....: $sAD_ParentOU - Parent OU where the new OU will be created (FQDN)
;                  $sAD_OU - OU to create in the the parent OU (Name without leading "OU=")
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_ParentOU does not exist
;                  |2 - $sAD_OU in $sAD_ParentOU already exists
;                  |3 - $sAD_OU is missing
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This does not create any attributes for the OU. Use function _AD_ModifyAttribute.
; Related .......: _AD_CreateUser, _AD_CreateGroup, _AD_AddUserToGroup, _AD_RemoveUserFromGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateOU($sAD_ParentOU, $sAD_OU)

	If Not _AD_ObjectExists($sAD_ParentOU, "distinguishedName") Then Return SetError(1, 0, 0)
	If _AD_ObjectExists("OU=" & $sAD_OU & "," & $sAD_ParentOU, "distinguishedName") Then Return SetError(2, 0, 0)
	If $sAD_OU = "" Then Return SetError(3, 0, 0)

	Local $oAD_ParentOU = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_ParentOU)
	Local $oAD_OU = $oAD_ParentOU.Create("organizationalUnit", "OU=" & $sAD_OU)
	$oAD_OU.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_CreateOU

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_CreateUser
; Description ...: Creates and activates a user in the specified OU.
; Syntax.........: _AD_CreateUser($sAD_OU, $sAD_User, $sAD_CN)
; Parameters ....: $sAD_OU - OU to create the user in. Form is "OU=sampleou,OU=sampleparent,DC=sampledomain1,DC=sampledomain2"
;                  $sAD_User - Username, form is SamAccountName without leading 'CN='
;                  $sAD_CN - Common Name (without CN=) or RDN (Relative Distinguished Name) like "Lastname Firstname"
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User already exists
;                  |2 - $sAD_OU does not exist
;                  |3 - $sAD_CN is missing
;                  |4 - $sAD_User is missing
;                  |5 - $sAD_User could not be created. @extended is set to the error returned by LDAP
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This function only sets sAMAccountName (= $sAD_User) and userPrincipalName (e.g. $sAD_User@microsoft.com).
;                  All other attributes have to be set using function _AD_ModifyAttribute
; Related .......: _AD_CreateOU, _AD_CreateGroup, _AD_AddUserToGroup, _AD_RemoveUserFromGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateUser($sAD_OU, $sAD_User, $sAD_CN)

	If _AD_ObjectExists($sAD_User) Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_OU, "distinguishedName") Then Return SetError(2, 0, 0)
	If $sAD_CN = "" Then Return SetError(3, 0, 0)
	$sAD_CN = _AD_FixSpecialChars($sAD_CN)
	If $sAD_User = "" Then Return SetError(4, 0, 0)
	Local $oAD_OU = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_OU)
	Local $oAD_User = $oAD_OU.Create("User", "CN=" & $sAD_CN)
	If @error Or Not IsObj($oAD_User) Then Return SetError(5, @error, 0)
	$oAD_User.sAMAccountName = $sAD_User
	$oAD_User.userPrincipalName = $sAD_User & "@" & StringTrimLeft(StringReplace($sAD_DNSDomain, ",DC=", "."), 3)
	$oAD_User.pwdLastSet = -1 ; Set password to not expired
	$oAD_User.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	$oAD_User.AccountDisabled = False ; Activate User
	$oAD_User.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_CreateUser

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SetPassword
; Description ...: Sets a user's password, or clears it if no password is passed to the function.
; Syntax.........: _AD_SetPassword($sAD_User[, $sAD_Password=""[, $iAD_Expired = 0]])
; Parameters ....: $sAD_User - User for which to set the password (FQDN or sAMAccountName)
;                  $sAD_Password - Optional: Password to be set for $sAD_User. If $sAD_Password is "" then the password will be cleared (default)
;                  $iAD_Expired - Optional: 1 = the password has to be changed at next logon (Default = 0)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: KenE
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetPassword($sAD_User, $sAD_Password = "", $iAD_Expired = 0)

	If Not _AD_ObjectExists($sAD_User) Then Return SetError(1, 0, 0)
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMACccountName provided
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_User)
	$oAD_User.SetPassword($sAD_Password)
	If $iAD_Expired Then $oAD_User.Put("pwdLastSet", 0)
	$oAD_User.SetInfo()
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_SetPassword

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_CreateGroup
; Description ...: Creates a group in the specified OU.
; Syntax.........: _AD_CreateGroup($sAD_OU, $sAD_Group[, $iAD_Type = $ADS_GROUP_TYPE_GLOBAL_SECURITY])
; Parameters ....: $sAD_OU - OU to create the group in. Form is "OU=sampleou,OU=sampleparent,DC=sampledomain1,DC=sampledomain2" (FQDN)
;                  $sAD_Group - Groupname, form is SamAccountName without leading 'CN='
;                  $iAD_Type - Optional: Group type (default = $ADS_GROUP_TYPE_GLOBAL_SECURITY). NOTE: Global security must be 'BitOr'ed with a scope.
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Group already exists
;                  |2 - $sAD_OU does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This function only sets sAMAccountName and grouptype. All other attributes have to be set using
;                  function _AD_ModifyAttribute
; Related .......: _AD_CreateOU, _AD_CreateUser, _AD_AddUserToGroup, _AD_RemoveUserFromGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateGroup($sAD_OU, $sAD_Group, $iAD_Type = $ADS_GROUP_TYPE_GLOBAL_SECURITY)

	If _AD_ObjectExists($sAD_Group) Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_OU, "distinguishedName") Then Return SetError(2, 0, 0)

	Local $sAD_CN = "CN=" & _AD_FixSpecialChars($sAD_Group)
	Local $oAD_OU = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_OU)
	Local $oAD_Group = $oAD_OU.Create("Group", $sAD_CN)
	Local $sAD_SamAccountName = StringReplace($sAD_Group, ",", "")
	$sAD_SamAccountName = StringReplace($sAD_SamAccountName, "#", "")
	$sAD_SamAccountName = StringReplace($sAD_SamAccountName, "/", "")
	$oAD_Group.sAMAccountName = $sAD_SamAccountName
	$oAD_Group.grouptype = $iAD_Type
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_CreateGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_AddUserToGroup
; Description ...: Adds a user or computer to the specified group.
; Syntax.........: _AD_AddUserToGroup($sAD_Group, $sAD_User)
; Parameters ....: $sAD_Group - Groupname (FQDN or sAMAccountName)
;                  $sAD_User - Username or computername to be added to the group (FQDN or sAMAccountName)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Group does not exist
;                  |2 - $sAD_User (user or computer) does not exist
;                  |3 - $sAD_User (user or computer) is already a member of $sAD_Group
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Works for both computers and groups. The sAMAccountname of a computer requires a trailing "$" before converting it to a FQDN.
; Related .......: _AD_CreateOU, _AD_CreateUser, _AD_CreateGroup, _AD_RemoveUserFromGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_AddUserToGroup($sAD_Group, $sAD_User)

	If Not _AD_ObjectExists($sAD_Group) Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_User) Then Return SetError(2, 0, 0)
	If _AD_IsMemberOf($sAD_Group, $sAD_User) Then Return SetError(3, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMACccountName provided
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMACccountName provided
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_User) ; Retrieve the COM Object for the user
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group) ; Retrieve the COM Object for the group
	$oAD_Group.Add($oAD_User.AdsPath)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_AddUserToGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_RemoveUserFromGroup
; Description ...: Removes a user or computer from the specified group.
; Syntax.........: _AD_RemoveUserFromGroup($sAD_Group, $sAD_User)
; Parameters ....: $sAD_Group - Groupname (FQDN or sAMAccountName)
;                  $sAD_User - Username (FQDN or sAMAccountName)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Group does not exist
;                  |2 - $sAD_User does not exist
;                  |3 - $sAD_User is not a member of $sAD_Group
;                  |x - Error returned by SetInfo Function(Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Works for computer objects as well. Remember that the sAMAccountname of a computer needs a trailing "$" before converting it to a FQDN.
; Related .......: _AD_CreateOU, _AD_CreateUser, _AD_CreateGroup, _AD_AddUserToGroup
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_RemoveUserFromGroup($sAD_Group, $sAD_User)

	If Not _AD_ObjectExists($sAD_Group) Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_User) Then Return SetError(2, 0, 0)
	If Not _AD_IsMemberOf($sAD_Group, $sAD_User) Then Return SetError(3, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMACccountName provided
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMACccountName provided
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_User) ; Retrieve the COM Object for the user
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group) ; Retrieve the COM Object for the group
	$oAD_Group.Remove($oAD_User.AdsPath)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_RemoveUserFromGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_CreateComputer
; Description ...: Creates and enables a computer account. A specific, authenticated user/group can then use this account to add his or her workstation to the domain.
; Syntax.........: _AD_CreateComputer($sAD_OU, $sAD_Computer, $sAD_User)
; Parameters ....: $sAD_OU - OU to create the computer in. Form is "OU=sampleou,OU=sampleparent,DC=sampledomain1,DC=sampledomain2" (FQDN)
;                  $sAD_Computer - Computername, form is SamAccountName without trailing "$"
;                  $sAD_User - User or group that will be allowed to add the computer to the domain (SamAccountName)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_OU does not exist
;                  |2 - $sAD_Computer already defined in $sAD_OU
;                  |3 - $sAD_User does not exist
;                  |x - Error returned by SetInfo Function(Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: By default, any authenticated user can create up to 10 computer accounts in the domain (machine account quota).
;                  (see: http://technet.microsoft.com/en-us/library/cc780195(WS.10).aspx)
;                  To create the Access Control List you need certain permissions. If this permissions are missing you might be able to add the
;                  computer to the domain but the function will exit with failure and the ACL is not set.
;+
;                  Creating a computer object in AD does not permit a user to join a computer to the domain.
;                  Certain permissions have to be granted so that the user has rights to modify the computer object.
;                  When you create a computer account using the ADUC snap-in you have the option to select a
;                  user or group to manage the computer object and join a computer to the domain using that object.
;+
;                  When you use that method, the following access control entries (ACEs) are added to the
;                  access control list (ACL) of the computer object:
;                  * List Contents, Read All Properties, Delete, Delete Subtree, Read Permissions, All
;                    Extended Rights (i.e., Allowed to Authenticate, Change Password, Send As, Receive As, Reset Password)
;                  * Write Property for description
;                  * Write Property for sAMAccountName
;                  * Write Property for displayName
;                  * Write Property for Logon Information
;                  * Write Property for Account Restrictions
;                  * Validate write to DNS host name
;                  * Validated write for service principal name
; Related .......: _AD_CreateOU, _AD_JoinDomain
; Link ..........: http://www.wisesoft.co.uk/scripts/vbscript_create_a_computer_account_for_a_specific_user.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateComputer($sAD_OU, $sAD_Computer, $sAD_User)

	If Not _AD_ObjectExists($sAD_OU) Then Return SetError(1, 0, 0)
	If _AD_ObjectExists("CN=" & $sAD_Computer & "," & $sAD_OU) Then Return SetError(2, 0, 0)
	If Not _AD_ObjectExists($sAD_User) Then Return SetError(3, 0, 0)
	If StringMid($sAD_OU, 3, 1) <> "=" Then $sAD_OU = _AD_SamAccountNameToFQDN($sAD_OU) ; sAMACccountName provided
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_User = _AD_FQDNToSamAccountName($sAD_User) ; FQDN provided
	Local $oAD_Container = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_OU)
	Local $oAD_Computer = $oAD_Container.Create("Computer", "cn=" & $sAD_Computer)
	$oAD_Computer.Put("sAMAccountName", $sAD_Computer & "$")
	$oAD_Computer.Put("userAccountControl", BitOR($ADS_UF_PASSWD_NOTREQD, $ADS_UF_WORKSTATION_TRUST_ACCOUNT))
	$oAD_Computer.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Local $oAD_SD = $oAD_Computer.Get("ntSecurityDescriptor")
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	Local $oAD_ACE1 = ObjCreate("AccessControlEntry")
	$oAD_ACE1.Trustee = $sAD_User
	$oAD_ACE1.AccessMask = $ADS_RIGHT_GENERIC_READ
	$oAD_ACE1.AceFlags = 0
	$oAD_ACE1.AceType = $ADS_ACETYPE_ACCESS_ALLOWED
	Local $oAD_ACE2 = ObjCreate("AccessControlEntry")
	$oAD_ACE2.Trustee = $sAD_User
	$oAD_ACE2.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
	$oAD_ACE2.AceFlags = 0
	$oAD_ACE2.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE2.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE2.ObjectType = $ALLOWED_TO_AUTHENTICATE
	Local $oAD_ACE3 = ObjCreate("AccessControlEntry")
	$oAD_ACE3.Trustee = $sAD_User
	$oAD_ACE3.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
	$oAD_ACE3.AceFlags = 0
	$oAD_ACE3.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE3.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE3.ObjectType = $RECEIVE_AS
	Local $oAD_ACE4 = ObjCreate("AccessControlEntry")
	$oAD_ACE4.Trustee = $sAD_User
	$oAD_ACE4.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
	$oAD_ACE4.AceFlags = 0
	$oAD_ACE4.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE4.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE4.ObjectType = $SEND_AS
	Local $oAD_ACE5 = ObjCreate("AccessControlEntry")
	$oAD_ACE5.Trustee = $sAD_User
	$oAD_ACE5.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
	$oAD_ACE5.AceFlags = 0
	$oAD_ACE5.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE5.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE5.ObjectType = $USER_CHANGE_PASSWORD
	Local $oAD_ACE6 = ObjCreate("AccessControlEntry")
	$oAD_ACE6.Trustee = $sAD_User
	$oAD_ACE6.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
	$oAD_ACE6.AceFlags = 0
	$oAD_ACE6.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE6.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE6.ObjectType = $USER_FORCE_CHANGE_PASSWORD
	Local $oAD_ACE7 = ObjCreate("AccessControlEntry")
	$oAD_ACE7.Trustee = $sAD_User
	$oAD_ACE7.AccessMask = $ADS_RIGHT_DS_WRITE_PROP
	$oAD_ACE7.AceFlags = 0
	$oAD_ACE7.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE7.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE7.ObjectType = $USER_ACCOUNT_RESTRICTIONS
	Local $oAD_ACE8 = ObjCreate("AccessControlEntry")
	$oAD_ACE8.Trustee = $sAD_User
	$oAD_ACE8.AccessMask = $ADS_RIGHT_DS_SELF
	$oAD_ACE8.AceFlags = 0
	$oAD_ACE8.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE8.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE8.ObjectType = $VALIDATED_DNS_HOST_NAME
	Local $oAD_ACE9 = ObjCreate("AccessControlEntry")
	$oAD_ACE9.Trustee = $sAD_User
	$oAD_ACE9.AccessMask = $ADS_RIGHT_DS_SELF
	$oAD_ACE9.AceFlags = 0
	$oAD_ACE9.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE9.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE9.ObjectType = $VALIDATED_SPN
	$oAD_DACL.AddAce($oAD_ACE1)
	$oAD_DACL.AddAce($oAD_ACE2)
	$oAD_DACL.AddAce($oAD_ACE3)
	$oAD_DACL.AddAce($oAD_ACE4)
	$oAD_DACL.AddAce($oAD_ACE5)
	$oAD_DACL.AddAce($oAD_ACE6)
	$oAD_DACL.AddAce($oAD_ACE7)
	$oAD_DACL.AddAce($oAD_ACE8)
	$oAD_DACL.AddAce($oAD_ACE9)
	$oAD_SD.DiscretionaryAcl = $oAD_DACL
	$oAD_Computer.Put("ntSecurityDescriptor", $oAD_SD)
	$oAD_Computer.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_CreateComputer

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ModifyAttribute
; Description ...: Modifies an attribute of the given object to the value specified.
; Syntax.........: _AD_ModifyAttribute($sAD_Object, $sAD_Attribute[, $sAD_Value = ""[, $iAD_Option = 1]])
; Parameters ....: $sAD_Object - Object (user, group ...) to add/delete/modify an attribute (sAMAccountName or FQDN)
;                  $sAD_Attribute - Attribute to add/delete/modify
;                  $sAD_Value - Optional: Value to modify the attribute to. Use a blank string ("") to delete the attribute (default).
;                  +$sAD_Value can be a single value (as a string) or a multi-value (as a one-dimensional array)
;                  $iAD_Option - Optional: Indicates the mode of modification: Append, Replace, Remove, and Delete
;                  |1 - CLEAR: remove all the property value(s) from the object (default when $sAD_value = "")
;                  |2 - UPDATE: replace the current value(s) with the specified value(s)
;                  |3 - APPEND: append the specified value(s) to the existing values(s)
;                  |4 - DELETE: delete the specified value(s) from the object
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_GetObjectAttribute, _AD_GetObjectProperties
; Link ..........: http://msdn.microsoft.com/en-us/library/aa746353(VS.85).aspx (ADS_PROPERTY_OPERATION_ENUM Enumeration)
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ModifyAttribute($sAD_Object, $sAD_Attribute, $sAD_Value = "", $iAD_Option = 1)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	$oAD_Object.GetInfo
	If $sAD_Value = "" Then
		$oAD_Object.PutEx(1, $sAD_Attribute, 0) ; CLEAR: remove all the property value(s) from the object
	ElseIf $iAD_Option = 3 Then
		$oAD_Object.PutEx(3, $sAD_Attribute, $sAD_Value) ; APPEND: append the specified value(s) to the existing values(s)
	ElseIf IsArray($sAD_Value) Then
		$oAD_Object.PutEx(2, $sAD_Attribute, $sAD_Value) ; UPDATE: replace the current value(s) with the specified value(s)
	Else
		$oAD_Object.Put($sAD_Attribute, $sAD_Value) ; sets the value(s) of an attribute
	EndIf
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_ModifyAttribute

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_RenameObject
; Description ...: Renames an object within an OU.
; Syntax.........: _AD_RenameObject($sAD_Object, $sAD_CN)
; Parameters ....: $sAD_Object - Object (user, group, computer) to rename (FQDN or sAMAccountName)
;                  $sAD_CN - New Name (relative name) of the object in the current OU without CN=
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by MoveHere function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Renames an object within the same OU. You can not move objects to another OU with this function.
; Related .......: _AD_MoveObject, _AD_DeleteObject
; Link ..........: http://msdn.microsoft.com/en-us/library/aa705991(v=VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_RenameObject($sAD_Object, $sAD_CN)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $oAD_OU = __AD_ObjGet($oAD_Object.Parent) ; Get the object of the OU/CN where the object resides
	$sAD_CN = "CN=" & _AD_FixSpecialChars($sAD_CN) ; escape all special characters
	$oAD_OU.MoveHere("LDAP://" & $sAD_HostServer & "/" & $sAD_Object, $sAD_CN)
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_RenameObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_MoveObject
; Description ...: Moves an object to another OU.
; Syntax.........: _AD_MoveObject($sAD_OU, $sAD_Object[, $sAD_DisplayName = ""])
; Parameters ....: $sAD_OU - Target OU for the object move (FQDN)
;                  $sAD_Object - Object (user, group, computer) to move (FQDN or sAMAccountName)
;                  $sAD_CN - Optional: New Name of the object in the target OU. Common Name or RDN (Relative Distinguished Name) like "Lastname Firstname" without leading "CN="
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_OU does not exist
;                  |2 - $sAD_Object does not exist
;                  |3 - Object already exists in the target OU
;                  |x - Error returned by MoveHere function (Missing permission etc.)
; Author ........: water
; Modified.......:
; Remarks .......: You must escape commas in $sAD_Object with a backslash. E.g. "CN=Lastname\, Firstname,OU=..."
; Related .......: _AD_RenameObject, _AD_DeleteObject
; Link ..........: http://msdn.microsoft.com/en-us/library/aa705991(v=VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_MoveObject($sAD_OU, $sAD_Object, $sAD_CN = "")

	If Not _AD_ObjectExists($sAD_OU, "distinguishedName") Then Return SetError(1, 0, 0)
	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(2, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	If $sAD_CN = "" Then
		$sAD_CN = "CN=" & _AD_FixSpecialChars(_AD_GetObjectAttribute($sAD_Object, "cn"))
	Else
		$sAD_CN = "CN=" & _AD_FixSpecialChars($sAD_CN) ; escape all special characters
	EndIf
	If _AD_ObjectExists($sAD_CN & "," & $sAD_OU, "distinguishedName") Then Return SetError(3, 0, 0)
	Local $oAD_OU = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_OU) ; Pointer to the destination container
	$oAD_OU.MoveHere("LDAP://" & $sAD_HostServer & "/" & $sAD_Object, $sAD_CN)
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_MoveObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_DeleteObject
; Description ...: Deletes the specified object.
; Syntax.........: _AD_DeleteObject($sAD_Object, $sAD_Class)
; Parameters ....: $sAD_Object - Object (user, group, computer etc.) to delete (FQDN or sAMAccountName)
;                  $sAD_Class - The schema class object to delete ("user", "computer", "group", "contact" etc). Can be derived using _AD_GetObjectClass().
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by Delete function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_RenameObject, _AD_MoveObject
; Link ..........: http://msdn.microsoft.com/en-us/library/aa705988(v=VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_DeleteObject($sAD_Object, $sAD_Class)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	Local $sAD_CN
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $oAD_OU = __AD_ObjGet($oAD_Object.Parent) ; Get the object of the OU/CN where the object resides
	If $sAD_Class = "organizationalUnit" Then
		$sAD_CN = "OU=" & _AD_GetObjectAttribute($sAD_Object, "ou")
	Else
		$sAD_CN = "CN=" & _AD_GetObjectAttribute($sAD_Object, "cn")
	EndIf
	$oAD_OU.Delete($sAD_Class, $sAD_CN)
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_DeleteObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SetAccountExpire
; Description ...: Modifies the specified user or computer account expiration date/time or sets the account to never expire.
; Syntax.........: _AD_SetAccountExpire($sAD_Object, $sAD_DateTime)
; Parameters ....: $sAD_Object - User or computer account to set expiration date/time (sAMAccountName or FQDN)
;                  $sAD_DateTime - Expiration date/time in format: yyyy-mm-dd hh:mm:ss (local time) or "01/01/1970" to never expire
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: KenE
; Modified.......: water
; Remarks .......: Use the following syntax for the date/time:
;                  01/01/1970 = never expire
;                  yyyy-mm-dd hh:mm:ss= "international format" - always works
;                  xx/xx/xx <time> = "localized format" - the format depends on the local date/time settings
;+
;                  Richard L. Mueller:
;                  "In Active Directory Users and Computers you can specify the date when a user account expires on the "Account"
;                  tab of the user properties dialog. This date is stored in the accountExpires attribute of the user object.
;                  There is also a property method called AccountExpirationDate, exposed by the IADsUser interface, that can be
;                  used to display and set this date. If you’ve ever compared accountExpires and AccountExpirationDate with the
;                  date shown in ADUC, you may have wondered what’s going on. It is common for the values to differ by a day,
;                  sometimes even two days."
; Related .......:
; Link ..........: http://www.rlmueller.net/AccountExpires.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetAccountExpire($sAD_Object, $sAD_DateTime)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	$oAD_Object.AccountExpirationDate = $sAD_DateTime
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_SetAccountExpire

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_DisablePasswordExpire
; Description ...: Modifies specified users password to not expire.
; Syntax.........: _AD_DisablePasswordExpire($sAD_Object)
; Parameters ....: $sAD_Object - User account to disable password expiration (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: KenE
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo, _AD_EnablePasswordExpire
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_DisablePasswordExpire($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $iAD_UAC = $oAD_Object.Get("userAccountControl")
	$oAD_Object.Put("userAccountControl", BitOR($iAD_UAC, $ADS_UF_DONT_EXPIRE_PASSWD))
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_DisablePasswordExpire

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_EnablePasswordExpire
; Description ...: Sets users password to expire.
; Syntax.........: _AD_EnablePasswordExpire($sAD_Object)
; Parameters ....: $sAD_Object - User account to enable password expiration (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Joe2010
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_EnablePasswordChange,  _AD_DisablePasswordChange, _AD_GetPasswordInfo, _AD_DisablePasswordExpire
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_EnablePasswordExpire($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $iAD_UAC = $oAD_Object.Get("userAccountControl")
	$oAD_Object.Put("userAccountControl", BitAND($iAD_UAC, BitNOT($ADS_UF_DONT_EXPIRE_PASSWD)))
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_EnablePasswordExpire

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_EnablePasswordChange
; Description ...: Disables the 'User Cannot Change Password' option, allowing the user to change their password.
; Syntax.........: _AD_EnablePasswordChange($sAD_Object)
; Parameters ....: $sAD_Object - User account to enable changing the password (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: KenE
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_DisablePasswordChange, _AD_GetPasswordInfo
; Link ..........: Example VBS see: http://gallery.technet.microsoft.com/ScriptCenter/en-us/ced14c6c-d16a-4cd8-b7d1-90d716c0445f or How to use Visual Basic and ADsSecurity.dll to properly order ACEs in an ACL. See: http://support.microsoft.com/kb/269159/en-us -
; Example .......: Yes
; ===============================================================================================================================
Func _AD_EnablePasswordChange($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $bAD_Self, $bAD_Everyone, $bAD_Modified, $sAD_Self = "NT AUTHORITY\SELF", $sAD_Everyone = "EVERYONE", $aAD_Temp
	; Get the language dependant well known accounts for SELF and EVERYONE
	$aAD_Temp = _Security__LookupAccountSid("S-1-5-10")
	If IsArray($aAD_Temp) Then $sAD_Self = $aAD_Temp[1] & "\" & $aAD_Temp[0]
	$aAD_Temp = _Security__LookupAccountSid("S-1-1-0")
	If IsArray($aAD_Temp) Then $sAD_Everyone = $aAD_Temp[0]
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $oAD_SD = $oAD_Object.Get("nTSecurityDescriptor")
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	; Search for ACE's for Change Password and modify
	$bAD_Self = False
	$bAD_Everyone = False
	$bAD_Modified = False
	For $oAD_ACE In $oAD_DACL
		If StringUpper($oAD_ACE.objectType) = StringUpper($USER_CHANGE_PASSWORD) Then
			If StringUpper($oAD_ACE.Trustee) = $sAD_Self Then
				If $oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT Then
					$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
					$bAD_Modified = True
				EndIf
				$bAD_Self = True
			EndIf
			If StringUpper($oAD_ACE.Trustee) = $sAD_Everyone Then
				If $oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT Then
					$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
					$bAD_Modified = True
				EndIf
				$bAD_Everyone = True
			EndIf
		EndIf
	Next
	; If ACE's found and modified, save changes and return
	If $bAD_Self And $bAD_Everyone Then
		If $bAD_Modified Then
			$oAD_SD.discretionaryACL = __AD_ReorderACE($oAD_DACL)
			$oAD_Object.Put("ntSecurityDescriptor", $oAD_SD)
			$oAD_Object.SetInfo
		EndIf
	Else
		; If ACE's not found, add to DACL
		If $bAD_Self = False Then
			; Create the ACE for Self
			Local $oAD_ACESelf = ObjCreate("AccessControlEntry")
			$oAD_ACESelf.Trustee = $sAD_Self
			$oAD_ACESelf.AceFlags = 0
			$oAD_ACESelf.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
			$oAD_ACESelf.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
			$oAD_ACESelf.objectType = $USER_CHANGE_PASSWORD
			$oAD_ACESelf.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
			$oAD_DACL.AddAce($oAD_ACESelf)
		EndIf
		If $bAD_Everyone = False Then
			; Create the ACE for Everyone
			Local $oAD_ACEEveryone = ObjCreate("AccessControlEntry")
			$oAD_ACEEveryone.Trustee = $sAD_Everyone
			$oAD_ACEEveryone.AceFlags = 0
			$oAD_ACEEveryone.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
			$oAD_ACEEveryone.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
			$oAD_ACEEveryone.objectType = $USER_CHANGE_PASSWORD
			$oAD_ACEEveryone.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
			$oAD_DACL.AddAce($oAD_ACEEveryone)
		EndIf
		$oAD_SD.discretionaryACL = __AD_ReorderACE($oAD_DACL)
		; Update the user object
		$oAD_Object.Put("ntSecurityDescriptor", $oAD_SD)
		$oAD_Object.SetInfo
	EndIf
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_EnablePasswordChange

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_DisablePasswordChange
; Description ...: Enables the 'User Cannot Change Password' option, preventing the user from changing their password.
; Syntax.........: _AD_DisablePasswordChange($sAD_Object)
; Parameters ....: $sAD_Object - User account to disallow changing his password (SAMAccountNmae or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: KenE
; Modified.......: water
; Remarks .......:
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange, _AD_GetPasswordInfo
; Link ..........: How to set the "User Cannot Change Password" option by using a program. See: http://support.microsoft.com/kb/301287/en-us -
; Example .......: Yes
; ===============================================================================================================================
Func _AD_DisablePasswordChange($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $bAD_Self, $bAD_Everyone, $bAD_Modified, $sAD_Self = "NT AUTHORITY\SELF", $sAD_Everyone = "EVERYONE", $aAD_Temp
	; Get the language dependant well known accounts for SELF and EVERYONE
	$aAD_Temp = _Security__LookupAccountSid("S-1-5-10")
	If IsArray($aAD_Temp) Then $sAD_Self = $aAD_Temp[1] & "\" & $aAD_Temp[0]
	$aAD_Temp = _Security__LookupAccountSid("S-1-1-0")
	If IsArray($aAD_Temp) Then $sAD_Everyone = $aAD_Temp[0]
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $oAD_SD = $oAD_Object.Get("nTSecurityDescriptor")
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	; Search for ACE's for Change Password and modify
	$bAD_Self = False
	$bAD_Everyone = False
	$bAD_Modified = False
	For $oAD_ACE In $oAD_DACL
		If StringUpper($oAD_ACE.objectType) = StringUpper($USER_CHANGE_PASSWORD) Then
			If StringUpper($oAD_ACE.Trustee) = $sAD_Self Then
				If $oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
					$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT
					$bAD_Modified = True
				EndIf
				$bAD_Self = True
			EndIf
			If StringUpper($oAD_ACE.Trustee) = $sAD_Everyone Then
				If $oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
					$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT
					$bAD_Modified = True
				EndIf
				$bAD_Everyone = True
			EndIf
		EndIf
	Next
	; If ACE's found and modified, save changes and return
	If $bAD_Self And $bAD_Everyone Then
		If $bAD_Modified Then
			$oAD_SD.discretionaryACL = __AD_ReorderACE($oAD_DACL)
			$oAD_Object.Put("ntSecurityDescriptor", $oAD_SD)
			$oAD_Object.SetInfo
		EndIf
	Else
		; If ACE's not found, add to DACL
		If $bAD_Self = False Then
			; Create the ACE for Self
			Local $oAD_ACESelf = ObjCreate("AccessControlEntry")
			$oAD_ACESelf.Trustee = $sAD_Self
			$oAD_ACESelf.AceFlags = 0
			$oAD_ACESelf.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT
			$oAD_ACESelf.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
			$oAD_ACESelf.objectType = $USER_CHANGE_PASSWORD
			$oAD_ACESelf.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
			$oAD_DACL.AddAce($oAD_ACESelf)
		EndIf
		If $bAD_Everyone = False Then
			; Create the ACE for Everyone
			Local $oAD_ACEEveryone = ObjCreate("AccessControlEntry")
			$oAD_ACEEveryone.Trustee = $sAD_Everyone
			$oAD_ACEEveryone.AceFlags = 0
			$oAD_ACEEveryone.AceType = $ADS_ACETYPE_ACCESS_DENIED_OBJECT
			$oAD_ACEEveryone.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
			$oAD_ACEEveryone.objectType = $USER_CHANGE_PASSWORD
			$oAD_ACEEveryone.AccessMask = $ADS_RIGHT_DS_CONTROL_ACCESS
			$oAD_DACL.AddAce($oAD_ACEEveryone)
		EndIf
		$oAD_SD.discretionaryACL = __AD_ReorderACE($oAD_DACL)
		; Update the user object
		$oAD_Object.Put("ntSecurityDescriptor", $oAD_SD)
		$oAD_Object.SetInfo
	EndIf
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_DisablePasswordChange

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_UnlockObject
; Description ...: Unlocks an AD object (user account, computer account).
; Syntax.........: _AD_UnlockObject($sAD_Object)
; Parameters ....: $sAD_Object - User account or computer account to unlock (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_IsObjectLocked, _AD_GetObjectsLocked
; Link ..........: http://www.rlmueller.net/Programs/IsUserLocked.txt
; Example .......: Yes
; ===============================================================================================================================
Func _AD_UnlockObject($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	$oAD_Object.IsAccountLocked = False
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_UnlockObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_DisableObject
; Description ...: Disables an AD object (user account, computer account).
; Syntax.........: _AD_DisableObject($sAD_Object)
; Parameters ....: $sAD_Object - User account or computer account to disable (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_IsObjectDisabled, _AD_EnableObject, _AD_GetObjectsDisabled
; Link ..........: http://www.wisesoft.co.uk/scripts/vbscript_enable-disable_user_account_1.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_DisableObject($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $iAD_UAC = $oAD_Object.Get("userAccountControl")
	$oAD_Object.Put("userAccountControl", BitOR($iAD_UAC, $ADS_UF_ACCOUNTDISABLE))
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_DisableObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_EnableObject
; Description ...: Enables an AD object (user account, computer account).
; Syntax.........: _AD_EnableObject($sAD_Object)
; Parameters ....: $sAD_Object - User account or computer account to enable (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Object does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_IsObjectDisabled, _AD_DisableObject, _AD_GetObjectsDisabled
; Link ..........: http://www.wisesoft.co.uk/scripts/vbscript_enable-disable_user_account_1.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_EnableObject($sAD_Object)

	If Not _AD_ObjectExists($sAD_Object) Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $iAD_UAC = $oAD_Object.Get("userAccountControl")
	$oAD_Object.Put("userAccountControl", BitAND($iAD_UAC, BitNOT($ADS_UF_ACCOUNTDISABLE)))
	$oAD_Object.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_EnableObject

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetPasswordInfo
; Description ...: Returns password information retrieved from the domain policy and the specified user or computer account.
; Syntax.........: _AD_GetPasswordInfo([$sAD_SamAccountName = @UserName])
; Parameters ....: $sAD_Object - Optional: User or computer account to get password info for (default = @UserName). Format is sAMAccountName or FQDN
; Return values .: Success - Returns a one-based array with the following information:
;                  |1 - Maximum Password Age (days)
;                  |2 - Minimum Password Age (days)
;                  |3 - Enforce Password History (# of passwords remembered)
;                  |4 - Minimum Password Length
;                  |5 - Account Lockout Duration (minutes). 0 means the account has to be unlocked manually by an administrator
;                  |6 - Account Lockout Threshold (invalid logon attempts)
;                  |7 - Reset account lockout counter after (minutes)
;                  |8 - Password last changed (YYYY/MM/DD HH:MM:SS in local time of the calling user) or "1601/01/01 00:00:00" (means "Password has never been set")
;                  |9 - Password expires (YYYY/MM/DD HH:MM:SS in local time of the calling user) or empty when password has not been set before or never expires
;                  |10 - Password last changed (YYYY/MM/DD HH:MM:SS in UTC) or "1601/01/01 00:00:00" (means "Password has never been set")
;                  |11 - Password expires (YYYY/MM/DD HH:MM:SS in UTC) or empty when password has not been set before or never expires
;                  |12 - Password properties. Part of Domain Policy. A bit field to indicate complexity / storage restrictions
;                  |      1 - DOMAIN_PASSWORD_COMPLEX
;                  |      2 - DOMAIN_PASSWORD_NO_ANON_CHANGE
;                  |      4 - DOMAIN_PASSWORD_NO_CLEAR_CHANGE
;                  |      8 - DOMAIN_LOCKOUT_ADMINS
;                  |     16 - DOMAIN_PASSWORD_STORE_CLEARTEXT
;                  |     32 - DOMAIN_REFUSE_PASSWORD_CHANGE
;                  Failure - "", sets @error to:
;                  |1 - $sAD_Object not found
;                  Warning - Returns a one-based array (see Success), sets @error to:
;                  |2 - Password does not expire (User Access Control - UAC - is set)
;                  |3 - Password has never been set
;                  |4 - The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire
;                  |The @error value can be a combination of the above values e.g. 5 = 2 (Password does not expire) + 3 (Password has never been set)
; Author ........: water
; Modified.......:
; Remarks .......: For details about password properties please check: http://msdn.microsoft.com/en-us/library/aa375371(v=vs.85).aspx
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange
; Link ..........: http://www.autoitscript.com/forum/index.php?showtopic=86247&view=findpost&p=619073, http://windowsitpro.com/article/articleid/81412/jsi-tip-8294-how-can-i-return-the-domain-password-policy-attributes.html
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetPasswordInfo($sAD_Object = @UserName)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, "")
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $iAD_Error = 0
	Local $aAD_PwdInfo[13] = [12]
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain)
	$aAD_PwdInfo[1] = Int(__AD_Int8ToSec($oAD_Object.Get("maxPwdAge"))) / 86400 ; Convert to Days
	$aAD_PwdInfo[2] = __AD_Int8ToSec($oAD_Object.Get("minPwdAge")) / 86400 ; Convert to Days
	$aAD_PwdInfo[3] = $oAD_Object.Get("pwdHistoryLength")
	$aAD_PwdInfo[4] = $oAD_Object.Get("minPwdLength")
	; Account lockout duration: http://msdn.microsoft.com/en-us/library/ms813429.aspx
	Local $oAD_Temp = $oAD_Object.Get("lockoutDuration")
	If $oAD_Temp.HighPart = 0x7FFFFFFF And $oAD_Temp.LowPart = 0xFFFFFFFF Then
		$aAD_PwdInfo[5] = 0 ; Account has to be unlocked manually by an admin
	Else
		$aAD_PwdInfo[5] = __AD_Int8ToSec($oAD_Temp) / 60 ; Convert to Minutes
	EndIf
	$aAD_PwdInfo[6] = $oAD_Object.Get("lockoutThreshold")
	$aAD_PwdInfo[7] = __AD_Int8ToSec($oAD_Object.Get("lockoutObservationWindow")) / 60 ; Convert to Minutes
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $sAD_PwdLastChanged = $oAD_User.Get("PwdLastSet")
	Local $iAD_UAC = $oAD_User.userAccountControl
	; Has user account password been changed before?
	If $sAD_PwdLastChanged.LowPart = 0 And $sAD_PwdLastChanged.HighPart = 0 Then
		$iAD_Error += 3
		$aAD_PwdInfo[8] = "1601/01/01 00:00:00"
		$aAD_PwdInfo[10] = "1601/01/01 00:00:00"
	Else
		Local $sAD_Temp = DllStructCreate("dword low;dword high")
		DllStructSetData($sAD_Temp, "Low", $sAD_PwdLastChanged.LowPart)
		DllStructSetData($sAD_Temp, "High", $sAD_PwdLastChanged.HighPart)
		; Have to convert to SystemTime because _Date_Time_FileTimeToStr has a bug (#1638)
		Local $sAD_Temp2 = _Date_Time_FileTimeToSystemTime(DllStructGetPtr($sAD_Temp))
		$aAD_PwdInfo[10] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp2, 1)
		; Convert PwdlastSet from UTC to Local Time
		$sAD_Temp2 = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sAD_Temp2))
		$aAD_PwdInfo[8] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp2, 1)
		; Is user account password set to expire?
		If BitAND($iAD_UAC, $ADS_UF_DONT_EXPIRE_PASSWD) = $ADS_UF_DONT_EXPIRE_PASSWD Or $aAD_PwdInfo[1] = 0 Then
			If BitAND($iAD_UAC, $ADS_UF_DONT_EXPIRE_PASSWD) = $ADS_UF_DONT_EXPIRE_PASSWD Then $iAD_Error += 2
			If $aAD_PwdInfo[1] = 0 Then $iAD_Error += 4 ; The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire
		Else
			$aAD_PwdInfo[11] = _DateAdd("d", $aAD_PwdInfo[1], $aAD_PwdInfo[10])
			$sAD_Temp2 = _Date_Time_EncodeSystemTime(StringMid($aAD_PwdInfo[11], 6, 2), StringMid($aAD_PwdInfo[11], 9, 2), StringMid($aAD_PwdInfo[11], 1, 4), StringMid($aAD_PwdInfo[11], 12, 2), StringMid($aAD_PwdInfo[11], 15, 2), StringMid($aAD_PwdInfo[11], 18, 2))
			; Convert PasswordExpires from UTC to Local Time
			$sAD_Temp2 = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sAD_Temp2))
			$aAD_PwdInfo[9] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp2, 1)
		EndIf
	EndIf
	$aAD_PwdInfo[12] = $oAD_Object.Get("pwdProperties")
	Return SetError($iAD_Error, 0, $aAD_PwdInfo)

EndFunc   ;==>_AD_GetPasswordInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListExchangeServers
; Description ...: Enumerates all Exchange Servers in the Forest.
; Syntax.........: _AD_ListExchangeServers()
; Parameters ....:
; Return values .: Success - Returns an one-based one dimensional array with the names of the Exchange Servers
;                  Failure - "", sets @error to:
;                  |1 - No Exchange Servers found. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_ListExchangeMailboxStores
; Link ..........: http://www.wisesoft.co.uk/scripts/vbscript_list_exchange_servers.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListExchangeServers()

	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Configuration & ">;(objectCategory=msExchExchangeServer);name;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	$oAD_RecordSet.MoveFirst
	Local $aAD_Result[1]
	Local $iCount1 = 1
	Do
		ReDim $aAD_Result[$iCount1 + 1]
		$aAD_Result[$iCount1] = $oAD_RecordSet.Fields("name").Value
		$oAD_RecordSet.MoveNext
		$iCount1 += 1
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	$aAD_Result[0] = UBound($aAD_Result, 1) - 1
	Return $aAD_Result

EndFunc   ;==>_AD_ListExchangeServers

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListExchangeMailboxStores
; Description ...: Enumerates all Exchange Mailbox Stores in the Forest.
; Syntax.........: _AD_ListExchangeMailboxStores()
; Parameters ....:
; Return values .: Success - Returns a one-based two dimensional array with the following information:
;                  |0 - name
;                  |1 - cn
;                  |2 - distinguishedName
;                  Failure - "", sets @error to:
;                  |1 - No Exchange Mailbox Stores found. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_ListExchangeServers
; Link ..........: http://www.wisesoft.co.uk/scripts/vbscript_list_exchange_mailbox_stores.aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListExchangeMailboxStores()

	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Configuration & ">;(objectCategory=msExchPrivateMDB);name,cn,distinguishedName;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	$oAD_RecordSet.MoveFirst
	Local $aAD_Result[1][3]
	Local $iCount1 = 1
	Do
		ReDim $aAD_Result[$iCount1 + 1][3]
		$aAD_Result[$iCount1][0] = $oAD_RecordSet.Fields("name").Value
		$aAD_Result[$iCount1][1] = $oAD_RecordSet.Fields("cn").Value
		$aAD_Result[$iCount1][2] = $oAD_RecordSet.Fields("distinguishedName").Value
		$oAD_RecordSet.MoveNext
		$iCount1 += 1
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	$aAD_Result[0][0] = UBound($aAD_Result, 1) - 1
	$aAD_Result[0][1] = UBound($aAD_Result, 2)
	Return $aAD_Result

EndFunc   ;==>_AD_ListExchangeMailboxStores

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetSystemInfo
; Description ...: Retrieves data describing the local computer if it is a member of a (at least) Windows 2000 domain.
; Syntax.........: _AD_GetSystemInfo()
; Parameters ....:
; Return values .: Success - Returns an one-based one dimensional array with the following information:
;                  |1 - distinguished name of the local computer
;                  |2 - DNS name of the local computer domain
;                  |3 - short name of the local computer domain
;                  |4 - DNS name of the local computer forest
;                  |5 - Local computer domain status: native mode (True) or mixed mode (False)
;                  |6 - distinguished name of the NTDS-DSA object for the DC that owns the primary domain controller role in the local computer domain
;                  |7 - distinguished name of the NTDS-DSA object for the DC that owns the schema role in the local computer forest
;                  |8 - site name of the local computer
;                  |9 - distinguished name of the current user
;                  Failure - "", sets @error to:
;                  |1 - Creation of object "ADSystemInfo" returned an error. See @extended
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/aa705962(VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetSystemInfo()

	Local $aAD_Result[10]
	Local $oAD_SystemInfo = ObjCreate("ADSystemInfo")
	If @error Or Not IsObj($oAD_SystemInfo) Then Return SetError(1, @error, "")
	$aAD_Result[1] = $oAD_SystemInfo.ComputerName
	$aAD_Result[2] = $oAD_SystemInfo.DomainDNSName
	$aAD_Result[3] = $oAD_SystemInfo.DomainShortName
	$aAD_Result[4] = $oAD_SystemInfo.ForestDNSName
	$aAD_Result[5] = $oAD_SystemInfo.IsNativeMode
	$aAD_Result[6] = $oAD_SystemInfo.PDCRoleOwner
	$aAD_Result[7] = $oAD_SystemInfo.SchemaRoleOwner
	$aAD_Result[8] = $oAD_SystemInfo.SiteName
	$aAD_Result[9] = $oAD_SystemInfo.UserName
	$aAD_Result[0] = 9
	Return $aAD_Result

EndFunc   ;==>_AD_GetSystemInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetManagedBy
; Description ...: Retrieves all groups that are managed by any or the specified user.
; Syntax.........: _AD_GetManagedBy([$sAD_ManagedBy = "*"])
; Parameters ....: $sAD_ManagedBy - Optional: Manager to filter the list of groups (default = *). Can be a SAMAccountname or a FQDN
; Return values .: Success - Returns a one-based two dimensional array with the following information:
;                  |0 - distinguished name of the group
;                  |1 - distinguished name of the manager for this group
;                  Failure - "", sets @error to:
;                  |1 - $sAD_ManagedBy is unknown
;                  |2 - No groups found. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......: This query returns all groups that have the attribute "managedBy" set or set to the specified manager.
;+
;                  To get a list of all groups that manager x manages (by querying just the user object) use:
;                    $Result = _AD_GetObjectAttribute("samAccountName of the manager","managedObjects")
;                    _ArrayDisplay($Result)
;+
;                  To return managers for OUs change "objectCategory=group" to "objectClass=organizationalUnit".
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetManagedBy($sAD_ManagedBy = "*")

	If $sAD_ManagedBy <> "*" Then
		If _AD_ObjectExists($sAD_ManagedBy) = 0 Then Return SetError(1, 0, "")
		If StringMid($sAD_ManagedBy, 3, 1) <> "=" Then $sAD_ManagedBy = _AD_SamAccountNameToFQDN($sAD_ManagedBy) ; sAMAccountName provided
	EndIf
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectCategory=group)(managedby=" & $sAD_ManagedBy & "))" & ";distinguishedName,managedby;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(2, @error, "")
	$oAD_RecordSet.MoveFirst
	Local $aAD_Result[1][2], $iCount1 = 1
	Do
		ReDim $aAD_Result[$iCount1 + 1][2]
		$aAD_Result[$iCount1][0] = $oAD_RecordSet.Fields("distinguishedName").Value
		$aAD_Result[$iCount1][1] = $oAD_RecordSet.Fields("managedBy").Value
		$oAD_RecordSet.MoveNext
		$iCount1 += 1
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	$aAD_Result[0][0] = UBound($aAD_Result, 1) - 1
	$aAD_Result[0][1] = UBound($aAD_Result, 2)
	Return $aAD_Result

EndFunc   ;==>_AD_GetManagedBy

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetManager
; Description ...: Retrieves all users that are managed by any or the specified user.
; Syntax.........: _AD_GetManager([$sAD_Manager = "*"])
; Parameters ....: $sAD_Manager - Optional: Manager to filter the list of users (default = *). Can be sAMAccountName or FQDN
; Return values .: Success - Returns a one-based two dimensional array with the following information:
;                  |0 - distinguished name of the user
;                  |1 - distinguished name of the manager for this user
;                  Failure - "", sets @error to:
;                  |1 - $sAD_Manager is unknown
;                  |2 - No users found. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......: This query returns all users that have the attribute "Manager" set or set to the specified manager.
;+
;                  To get a list of all users that manager x manages (by querying just the user object) use:
;                    $Result = _AD_GetObjectAttribute("samAccountName of the manager","directReports")
;                    _ArrayDisplay($Result)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetManager($sAD_Manager = "*")

	If $sAD_Manager <> "*" Then
		If _AD_ObjectExists($sAD_Manager) = 0 Then Return SetError(1, 0, "")
		If StringMid($sAD_Manager, 3, 1) <> "=" Then $sAD_Manager = _AD_SamAccountNameToFQDN($sAD_Manager) ; sAMAccountName provided
	EndIf
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectCategory=user)(manager=" & $sAD_Manager & "))" & ";distinguishedName,Manager;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(2, @error, "")
	$oAD_RecordSet.MoveFirst
	Local $aAD_Result[1][2], $iCount1 = 1
	Do
		ReDim $aAD_Result[$iCount1 + 1][2]
		$aAD_Result[$iCount1][0] = $oAD_RecordSet.Fields("distinguishedName").Value
		$aAD_Result[$iCount1][1] = $oAD_RecordSet.Fields("Manager").Value
		$oAD_RecordSet.MoveNext
		$iCount1 += 1
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	$aAD_Result[0][0] = UBound($aAD_Result, 1) - 1
	$aAD_Result[0][1] = UBound($aAD_Result, 2)
	Return $aAD_Result

EndFunc   ;==>_AD_GetManager

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetGroupAdmins
; Description ...: Returns an array of the administrator sAMAccountNames for the specified group (not including the group owner/manager).
; Syntax.........: _AD_GetGroupAdmins($sAD_Object)
; Parameters ....: $sAD_Object - group name. Can be SAMaccountName or FQDN
; Return values .: Success - Returns an one-based one dimensional array with the sAMAccountNames of the administrators for the specified group
;                  +(not including the group owner/manager)
;                  Failure - "", sets @error to:
;                  |1 - Group could not be found
;                  |2 - No administrators found
; Author ........: John Clelland
; Modified.......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetGroupAdmins($sAD_Object)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, "")
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $sAD_ManagedBy = $oAD_Group.Get("managedBy")
	Local $oAD_SD = $oAD_Group.Get("ntSecurityDescriptor")
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	Local $aAD_Admins[1] = [0], $sAD_samID, $sAD_ManagedBy_samID
	For $oAD_ACE In $oAD_DACL
		If $oAD_ACE.ObjectType = $SELF_MEMBERSHIP And $oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT And _
				BitAND($oAD_ACE.AccessMask, $ADS_RIGHT_DS_WRITE_PROP) = $ADS_RIGHT_DS_WRITE_PROP Then
			$sAD_samID = StringTrimLeft($oAD_ACE.Trustee, StringInStr($oAD_ACE.Trustee, "\"))
			; S-1-5-21: SECURITY_NT_NON_UNIQUE - See: http://msdn.microsoft.com/en-us/library/aa379649(VS.85).aspx
			If Not StringInStr($sAD_samID, "S-1-5-21") And Not StringInStr($sAD_samID, "Account Operator") Then _ArrayAdd($aAD_Admins, $sAD_samID)
		EndIf
	Next
	If $sAD_ManagedBy <> "" Then
		$sAD_ManagedBy_samID = _AD_FQDNToSamAccountName($sAD_ManagedBy)
		Local $iCount1
		Local $iAD_Owner = -1
		For $iCount1 = 1 To UBound($aAD_Admins) - 1
			If $aAD_Admins[$iCount1] = $sAD_ManagedBy_samID Then $iAD_Owner = $iCount1
		Next
		If $iAD_Owner <> -1 Then
			_ArrayDelete($aAD_Admins, $iAD_Owner)
		EndIf
	EndIf
	$aAD_Admins[0] = UBound($aAD_Admins) - 1
	If $aAD_Admins[0] = 0 Then Return SetError(2, 0, "")
	Return $aAD_Admins

EndFunc   ;==>_AD_GetGroupAdmins

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GroupManagerCanModify
; Description ...: Returns 1 if the manager of the group can modify the member list.
; Syntax.........: _AD_GroupManagerCanModify($sAD_Object)
; Parameters ....: $sAD_Object - FQDN of the group
; Return values .: Success - 1, Specified user can modify the member list
;                  Failure - 0, @error set
;                  |1 - Group does not exist
;                  |2 - The group manager can not modify the member list
;                  |3 - There is no manager assigned to the group
; Author ........: John Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_GroupAssignManager, _AD_GroupRemoveManager, _AD_SetGroupManagerCanModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GroupManagerCanModify($sAD_Object)

	If _AD_ObjectExists($sAD_Object) = 0 Then Return SetError(1, 0, 0)
	If StringMid($sAD_Object, 3, 1) <> "=" Then $sAD_Object = _AD_SamAccountNameToFQDN($sAD_Object) ; sAMAccountName provided
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Object)
	Local $sAD_ManagedBy = $oAD_Group.Get("managedBy")
	If $sAD_ManagedBy = "" Then Return SetError(3, 0, 0)
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_ManagedBy)
	Local $aAD_UserFQDN = StringSplit($sAD_ManagedBy, "DC=", 1)
	Local $sAD_Domain = StringTrimRight($aAD_UserFQDN[2], 1)
	Local $sAD_SamAccountName = $oAD_User.Get("sAMAccountName")
	Local $oAD_SD = $oAD_Group.Get("ntSecurityDescriptor")
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	Local $bAD_Match = False
	For $oAD_ACE In $oAD_DACL
		If StringLower($oAD_ACE.Trustee) = StringLower($sAD_Domain & "\" & $sAD_SamAccountName) And _
				$oAD_ACE.ObjectType = $SELF_MEMBERSHIP And _
				$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT And _
				BitAND($oAD_ACE.AccessMask, $ADS_RIGHT_DS_WRITE_PROP) = $ADS_RIGHT_DS_WRITE_PROP Then $bAD_Match = True
	Next
	If $bAD_Match Then
		Return 1
	Else
		Return SetError(2, 0, 0)
	EndIf

EndFunc   ;==>_AD_GroupManagerCanModify

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListPrintQueues
; Description ...: Enumerates all PrintQueues in the AD tree or for the specified spool server.
; Syntax.........: _AD_ListPrintQueues([$sAD_Servername=*])
; Parameters ....: $sAD_Servername - Optional: Short name of the spool server to process
; Return values .: Success - One-based two dimensional array with the following information:
;                  |0 - PrinterName: Short name of the PrintQueue
;                  |1 - ServerName: SpoolServerName.Domain
;                  |2 - DistinguishedName: FQDN of the PrintQueue
;                  Failure - "", @error set
;                  |1 - There is no PrintQueue available. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......: To get more (including multi-valued) attributes of a printqueue use _AD_GetObjectProperties
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/aa706091(VS.85).aspx, http://www.activxperts.com/activmonitor/windowsmanagement/scripts/printing/printerport/#LAPP.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListPrintQueues($sAD_Servername = "*")

	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(&(objectclass=printQueue)(shortservername=" & $sAD_Servername & "));distinguishedName,PrinterName,ServerName;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	Local $aAD_PrinterList[$oAD_RecordSet.RecordCount + 1][3] = [[0, 3]]
	$oAD_RecordSet.MoveFirst
	Do
		$aAD_PrinterList[0][0] += 1
		$aAD_PrinterList[$aAD_PrinterList[0][0]][0] = $oAD_RecordSet.Fields("printerName").Value
		$aAD_PrinterList[$aAD_PrinterList[0][0]][1] = $oAD_RecordSet.Fields("serverName").Value
		$aAD_PrinterList[$aAD_PrinterList[0][0]][2] = $oAD_RecordSet.Fields("distinguishedName").Value
		$oAD_RecordSet.MoveNext
	Until $oAD_RecordSet.EOF
	$oAD_RecordSet.Close
	Return $aAD_PrinterList

EndFunc   ;==>_AD_ListPrintQueues

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SetGroupManagerCanModify
; Description ...: Sets the Group manager to be able to modify the member list.
; Syntax.........: _AD_SetGroupManagerCanModify($sAD_Group)
; Parameters ....: $sAD_Group - Groupname (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, @error set
;                  |1 - $sAD_Group does not exist
;                  |2 - Group manager can already modify the member list
;                  |3 - $sAD_Group has no manager assigned
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_GroupAssignManager, _AD_GroupManagerCanModify, _AD_GroupRemoveManager
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetGroupManagerCanModify($sAD_Group)

	Local Const $ADS_OPTION_SECURITY_MASK = 3 ; See: http://msdn.microsoft.com/en-us/library/windows/desktop/aa772273(v=vs.85).aspx
	Local Const $ADS_SECURITY_INFO_DACL = 0x4 ; See: http://msdn.microsoft.com/en-us/library/windows/desktop/aa772293(v=vs.85).aspx
	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	If _AD_GroupManagerCanModify($sAD_Group) = 1 Then Return SetError(2, 0, 0)
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group)
	Local $sAD_ManagedBy = $oAD_Group.Get("managedBy")
	If $sAD_ManagedBy = "" Then Return SetError(3, 0, 0)
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_ManagedBy)
	Local $aAD_UserFQDN = StringSplit($sAD_ManagedBy, "DC=", 1)
	Local $sAD_Domain = StringTrimRight($aAD_UserFQDN[2], 1)
	Local $sAD_SamAccountName = $oAD_User.Get("sAMAccountName")
	Local $oAD_SD = $oAD_Group.Get("ntSecurityDescriptor")
	$oAD_SD.Owner = $sAD_Domain & "\" & @UserName
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	Local $oAD_ACE = ObjCreate("AccessControlEntry")
	$oAD_ACE.Trustee = $sAD_Domain & "\" & $sAD_SamAccountName
	$oAD_ACE.AccessMask = $ADS_RIGHT_DS_WRITE_PROP
	$oAD_ACE.AceFlags = 0
	$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
	$oAD_ACE.Flags = $ADS_FLAG_OBJECT_TYPE_PRESENT
	$oAD_ACE.ObjectType = $SELF_MEMBERSHIP
	$oAD_DACL.AddAce($oAD_ACE)
	$oAD_SD.DiscretionaryAcl = __AD_ReorderACE($oAD_DACL)
	; Set options to only update the DACL. See: http://www.activedir.org/ListArchives/tabid/55/view/topic/postid/28231/Default.aspx
	$oAD_Group.Setoption($ADS_OPTION_SECURITY_MASK, $ADS_SECURITY_INFO_DACL)
	$oAD_Group.Put("ntSecurityDescriptor", $oAD_SD)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_SetGroupManagerCanModify

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GroupAssignManager
; Description ...: Assigns the user as group manager.
; Syntax.........: _AD_GroupAssignManager($sAD_Group, $sAD_User)
; Parameters ....: $sAD_Group - Groupname (sAMAccountName or FQDN)
;                  $sAD_User - User to assign as manager (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, @error set
;                  |1 - $sAD_Group does not exist
;                  |2 - $sAD_User does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_SetGroupManagerCanModify, _AD_GroupManagerCanModify, _AD_GroupRemoveManager
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GroupAssignManager($sAD_Group, $sAD_User)

	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, 0)
	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(2, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMAccountName provided
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group)
	$oAD_Group.Put("managedBy", $sAD_User)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_GroupAssignManager

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GroupRemoveManager
; Description ...: Remove the group manager from a group or only remove the manager's modify permission.
; Syntax.........: _AD_GroupRemoveManager($sAD_Group[, $bAD_Flag = False])
; Parameters ....: $sAD_Group - Groupname (sAMAccountName or FQDN)
;                  $bAD_Flag  - Optional: if True the function only removes the manager's modify permission
; Return values .: Success - 1
;                  Failure - 0, @error set
;                  |1 - $sAD_Group does not exist
;                  |2 - $sAD_Group does not have a manager assigned
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......: _AD_SetGroupManagerCanModify, _AD_GroupManagerCanModify, _AD_GroupAssignManager
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GroupRemoveManager($sAD_Group, $bAD_Flag = False)

	If _AD_ObjectExists($sAD_Group) = 0 Then Return SetError(1, 0, 0)
	If StringMid($sAD_Group, 3, 1) <> "=" Then $sAD_Group = _AD_SamAccountNameToFQDN($sAD_Group) ; sAMAccountName provided
	Local $oAD_Group = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_Group)
	Local $sAD_ManagedBy = $oAD_Group.Get("managedBy")
	If $sAD_ManagedBy = "" Then Return SetError(2, 0, 0)
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_ManagedBy)
	Local $aAD_UserFQDN = StringSplit($sAD_ManagedBy, "DC=", 1)
	Local $sAD_Domain = StringTrimRight($aAD_UserFQDN[2], 1)
	Local $sAD_SamAccountName = $oAD_User.Get("sAMAccountName")
	Local $oAD_SD = $oAD_Group.Get("ntSecurityDescriptor")
	$oAD_SD.Owner = $sAD_Domain & "\" & @UserName
	Local $oAD_DACL = $oAD_SD.DiscretionaryAcl
	For $oAD_ACE In $oAD_DACL
		If StringLower($oAD_ACE.Trustee) = StringLower($sAD_Domain & "\" & $sAD_SamAccountName) And _
				$oAD_ACE.ObjectType = $SELF_MEMBERSHIP And _
				$oAD_ACE.AceType = $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT And _
				$oAD_ACE.AccessMask = $ADS_RIGHT_DS_WRITE_PROP Then _
				$oAD_DACL.RemoveAce($oAD_ACE)
	Next
	$oAD_SD.DiscretionaryAcl = $oAD_DACL
	$oAD_Group.Put("ntSecurityDescriptor", $oAD_SD)
	If Not $bAD_Flag Then $oAD_Group.PutEx(1, "managedBy", 0)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_GroupRemoveManager

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_AddEmailAddress
; Description ...: Appends an SMTP email address to the 'Email Addresses' tab of an Exchange-enabled AD account.
; Syntax.........: _AD_AddEmailAddress($sAD_User, $sAD_NewEmail[, $bAD_Primary = False])
; Parameters ....: $sAD_User - User account (sAMAccountName or FQDN)
;                  $sAD_NewEmail - Email address to add to the account
;                  $bAD_Primary - Optional: if True the new email address will be set as primary address
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |2 - Could not connect to $sAD_User. @extended is set to the error returned by LDAP
;                  |x - Error returned by GetEx or SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_AddEmailAddress($sAD_User, $sAD_NewEmail, $bAD_Primary = False)

	If _AD_ObjectExists($sAD_User) = 0 Then Return SetError(1, 0, 0)
	If StringMid($sAD_User, 3, 1) <> "=" Then $sAD_User = _AD_SamAccountNameToFQDN($sAD_User) ; sAMAccountName provided
	Local $oAD_User = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_User)
	If @error Or Not IsObj($oAD_User) Then Return SetError(2, @error, 0)
	Local $aAD_ProxyAddresses = $oAD_User.GetEx("proxyaddresses")
	If @error <> 0 Then Return SetError(@error, 0, 0)
	If $bAD_Primary Then
		For $iCount1 = 0 To UBound($aAD_ProxyAddresses) - 1
			If StringInStr($aAD_ProxyAddresses[$iCount1], "SMTP:", 1) Then
				$aAD_ProxyAddresses[$iCount1] = StringReplace($aAD_ProxyAddresses[$iCount1], "SMTP:", "smtp:", 0, 1)
			EndIf
		Next
		_ArrayAdd($aAD_ProxyAddresses, "SMTP:" & $sAD_NewEmail)
		$oAD_User.Put("mail", $sAD_NewEmail)
	Else
		_ArrayAdd($aAD_ProxyAddresses, "smtp:" & $sAD_NewEmail)
	EndIf
	$oAD_User.PutEx(2, "proxyaddresses", $aAD_ProxyAddresses)
	$oAD_User.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_AddEmailAddress

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_JoinDomain
; Description ...: Joins the computer to an already created computer account in a domain.
; Syntax.........: _AD_JoinDomain($sAD_Computer[, $sAD_UserParam = "", $sAD_PasswordParam = ""])
; Parameters ....: $sAD_Computer - Computername to join to the domain
;                  $sAD_UserParam - Optional: User with admin rights to join the computer to the domain (NetBIOSName (domain\user) or user principal name (user@domain))
;                  +(Default = credentials under which the script is run or from _AD_Open are used)
;                  $sAD_PasswordParam - Optional: Password for $sAD_UserParam. (Default = credentials under which the script is run are used)
; Return values .: Success - 1
;                  Failure - 0, @error set
;                  |1 - $sAD_Computer account does not exist in the domain
;                  |2 - $sAD_UserParam does not exist in the domain
;                  |3 - WMI object could not be created. See @extended for error code. See remarks for further information
;                  |4 - The computer is already a member of the domain
;                  |5 - Joining the domain was not successful. @extended holds the Win32 error code (see: http://msdn.microsoft.com/en-us/library/ms681381(v=VS.85).aspx)
; Author ........: water
; Modified.......:
; Remarks .......: The domain to which the computer is joined to is the domain the user logged on to using AD_Open.
;                  If no credentials are passed to this function but have been used with _AD_Open() then the _AD_Open credentials will be used for this function.
;                  You have to reboot the computer after a successful join to the domain.
;                  The JoinDomainOrWorkgroup method is available only on Windows XP computer and Windows Server 2003 or later.
; Related .......: _AD_CreateComputer
; Link ..........: http://technet.microsoft.com/en-us/library/ee692588.aspx, http://msdn.microsoft.com/en-us/library/aa392154(VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_JoinDomain($sAD_Computer, $sAD_UserParam = "", $sAD_PasswordParam = "")

	If _AD_ObjectExists($sAD_Computer & "$") = 0 Then Return SetError(1, 0, 0)
	If $sAD_UserParam <> "" And _AD_ObjectExists($sAD_UserParam) = 0 Then Return SetError(2, 0, 0)
	Local $iAD_Result
	Local $sAD_DomainName = StringReplace(StringReplace($sAD_DNSDomain, "DC=", ""), ",", ".")
	; Create WMI object
	Local $oAD_Computer = ObjGet("winmgmts:{impersonationLevel=Impersonate}!\\" & $sAD_Computer & "\root\cimv2:Win32_ComputerSystem.Name='" & $sAD_Computer & "'")
	If @error Or Not IsObj($oAD_Computer) Then Return SetError(3, @error, 0)
	If $oAD_Computer.Domain = $sAD_DomainName Then Return SetError(4, 0, 0)
	; Join domain. Use credentials passed as parameters to this function or those used at _AD_Open, if any
	If $sAD_UserParam <> "" Then
		$iAD_Result = $oAD_Computer.JoinDomainOrWorkGroup($sAD_DomainName, $sAD_PasswordParam, $sAD_DomainName & "\" & $sAD_UserParam, Default, 1)
	ElseIf $sAD_UserId <> "" Then
		; Domain NetBIOS name and user account (domain\user) or user principal name (user@domain) is supported
		Local $sAD_Temp = $sAD_UserId
		If StringInStr($sAD_UserId, "\") = 0 And StringInStr($sAD_UserId, "@") = 0 Then $sAD_Temp = $sAD_DomainName & "\" & $sAD_UserId
		$iAD_Result = $oAD_Computer.JoinDomainOrWorkGroup($sAD_DomainName, $sAD_Password, $sAD_DomainName & "\" & $sAD_Temp, Default, 1)
	Else
		$iAD_Result = $oAD_Computer.JoinDomainOrWorkGroup($sAD_DomainName, Default, Default, Default, 1)
	EndIf
	If $iAD_Result <> 0 Then Return SetError(5, $iAD_Result, 0)
	Return 1

EndFunc   ;==>_AD_JoinDomain

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_UnJoinDomain
; Description ...: Unjoins the computer from its current domain and disables the computer account.
; Syntax.........: _AD_UnJoinDomain($sAD_Computer[, $sAD_Workgroup [,$sAD_UserParam, $sAD_PasswordParam]])
; Parameters ....: $sAD_Computer - Computername to unjoin from the domain
;                  $sAD_Workgroup - Optional: Workgroup the unjoined computer is assigned to (Default = Domain the computer was unjoined from)
;                  $sAD_UserParam - Optional: User with admin rights to unjoin the computer from the domain (NetBIOSName)
;                  +(Default = credentials under which the script is run or from _AD_Open are used)
;                  $sAD_PasswordParam - Optional: Password for $sAD_UserParam. (Default = credentials under which the script is run are used)
; Return values .: Success - 1
;                  Failure - 0, @error set
;                  |1 - $sAD_Computer account does not exist in the domain
;                  |2 - $sAD_UserParam does not exist in the domain
;                  |3 - WMI object could not be created. See @extended for error code. See remarks for further information
;                  |4 - The computer is a member of another or no domain
;                  |5 - Unjoining the domain was not successful. See @extended for error code. See remarks for further information
;                  |6 - Joining the Computer to the specified workgroup was not successful. See @extended for error code
; Author ........: water
; Modified.......:
; Remarks .......: The domain from which the computer is unjoined from has to be the domain the user logged on to using AD_Open.
;                  If no credentials are passed to this function but have been used with _AD_Open() then the _AD_Open credentials will be used for this function.
;                  If no workgroup is specified then the computer is assigned to a workgroup named like the domain the computer was unjoined from.
;                  You have to reboot the computer after a successful unjoin from the domain.
;                  The UnjoinDomainOrWorkgroup method is available only on Windows XP computer and Windows Server 2003 or later.
; Related .......:
; Link ..........: http://gallery.technet.microsoft.com/ScriptCenter/en-us/c2025ace-cb51-4136-9de9-db8871f79f62, http://technet.microsoft.com/en-us/library/ee692588.aspx, http://msdn.microsoft.com/en-us/library/aa393942(VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_UnJoinDomain($sAD_Computer, $sAD_Workgroup = "", $sAD_UserParam = "", $sAD_PasswordParam = "")

	Local $NETSETUP_ACCT_DELETE = 4 ; According to MS it should be 2 but only 4 works
	If _AD_ObjectExists($sAD_Computer & "$") = 0 Then Return SetError(1, 0, 0)
	If $sAD_UserParam <> "" And _AD_ObjectExists($sAD_UserParam) = 0 Then Return SetError(2, 0, 0)
	Local $iAD_Result
	Local $sAD_DomainName = StringReplace(StringReplace($sAD_DNSDomain, "DC=", ""), ",", ".")
	; Create WMI object
	Local $oAD_Computer = ObjGet("winmgmts:{impersonationLevel=Impersonate}!\\" & $sAD_Computer & "\root\cimv2:Win32_ComputerSystem.Name='" & $sAD_Computer & "'")
	If @error Or Not IsObj($oAD_Computer) Then Return SetError(3, @error, 0)
	If $oAD_Computer.Domain <> $sAD_DomainName Then Return SetError(4, 0, 0)
	; UnJoin domain. Use credentials passed as parameters to this function or those used at _AD_Open, if any
	If $sAD_UserParam <> "" Then
		$iAD_Result = $oAD_Computer.UnjoinDomainOrWorkGroup($sAD_PasswordParam, $sAD_DomainName & "\" & $sAD_UserParam, $NETSETUP_ACCT_DELETE)
	ElseIf $sAD_UserId <> "" Then
		; Domain NetBIOS name and user account (domain\user) or user principal name (user@domain) is supported
		Local $sAD_Temp = $sAD_UserId
		If StringInStr($sAD_UserId, "\") = 0 And StringInStr($sAD_UserId, "@") = 0 Then $sAD_Temp = $sAD_DomainName & "\" & $sAD_UserId
		$iAD_Result = $oAD_Computer.UnjoinDomainOrWorkGroup($sAD_Password, $sAD_Temp, $NETSETUP_ACCT_DELETE)
	Else
		$iAD_Result = $oAD_Computer.UnjoinDomainOrWorkGroup(Default, Default, $NETSETUP_ACCT_DELETE)
	EndIf
	If $iAD_Result <> 0 Then Return SetError(5, $iAD_Result, 0)
	; Move unjoined computer to another workgroup
	If $sAD_Workgroup <> "" Then
		$iAD_Result = $oAD_Computer.JoinDomainOrWorkGroup($sAD_Workgroup, Default, Default, Default, Default)
		If $iAD_Result <> 0 Then Return SetError(6, $iAD_Result, 0)
	EndIf
	Return 1

EndFunc   ;==>_AD_UnJoinDomain

; #FUNCTION#====================================================================================================================
; Name...........: _AD_SetPasswordExpire
; Description ...: Sets user's password as expired or not expired.
; Syntax.........: _AD_SetPasswordExpire($sAD_User[, $iAD_Flag = 0])
; Parameters ....: $sAD_User - User account (sAMAccountName or FQDN)
;                  $iAD_Flag - Optional: Sets the user's password as expired ($iAD_Flag = 0) or as not expired ($iAD_Flag = -1) (Default = 0)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |2 - $iAD_Flag has an invalid value
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Ethan Turk
; Modified.......:
; Remarks .......: When the user's password is set to expired he has to change the password at next logon
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetPasswordExpire($sAD_User, $iAD_Flag = 0)

	If Not _AD_ObjectExists($sAD_User) Then Return SetError(1, 0, 0)
	If $iAD_Flag <> 0 And $iAD_Flag <> -1 Then Return SetError(2, 0, 0)
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_User & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	$oAD_Object.Put("pwdLastSet", $iAD_Flag)
	$oAD_Object.SetInfo()
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_SetPasswordExpire

; #FUNCTION#====================================================================================================================
; Name...........: _AD_CreateMailbox
; Description ...: Creates a mailbox for a user.
; Syntax.........: _AD_CreateMailbox($sAD_User, $sAD_StoreName[, $sAD_Store[, $sAD_EMailServer[, $sAD_AdminGroup[, $sAD_EmailDomain]]]])
; Parameters ....: $sAD_User - User account (sAMAccountName or FQDN) to add mailbox to
;                  $sAD_StoreName - Mailbox storename
;                  $sAD_Store - Optional: Information store (Default = "First Storage Group")
;                  $sAD_EmailServer - Optional: Email server (Default = First server returned by _AD_ListExchangeServers())
;                  $sAD_AdminGroup - Optional: Administrative group in Exchange (Default= "First Administrative Group")
;                  $sAD_EmailDomain - Optional: Exchange Server Group name e.g. "My Company" (Default = Text after first DC= in $sAD_DNSDomain)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_User already has a mailbox
;                  |3 - _AD_ListExchangeServers() returned an error. Please see @extended for _AD_ListExchangeServers() return code
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: The mailbox is created using CDOEXM. For this function to work the Exchange administration tools have to be installed on the
;                  computer running the script.
;                  To set rights on the mailbox you have to run at least Exchange 2000 SP2.
;+
;                  If the Exchange administration tools are not installed on the PC running the script you could use an ADSI only solution.
;                  Set the mailNickname and displayName properties of the user and at least one of this: homeMTA, homeMDB or msExchHomeServerName and
;                  the RUS (Recipient Update Service) of Exchange 2000/2003 will create the mailbox for you.
;                  Be aware that this no longer works for Exchange 2007 and later.
; Related .......:
; Link ..........: http://www.msxfaq.de/code/makeuser.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateMailbox($sAD_User, $sAD_StoreName, $sAD_Store = "First Storage Group", $sAD_EMailServer = "", $sAD_AdminGroup = "First Administrative Group", $sAD_EmailDomain = "")

	If Not _AD_ObjectExists($sAD_User) Then Return SetError(1, 0, 0)
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_User & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_User = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	If $oAD_User.HomeMDB <> "" Then Return SetError(2, 0, 0)
	Local $aAD_Temp
	If $sAD_EmailDomain = "" Then
		$aAD_Temp = StringSplit($sAD_DNSDomain, ",")
		$sAD_EmailDomain = StringMid($aAD_Temp[1], 4)
	EndIf
	If $sAD_EMailServer = "" Then
		$aAD_Temp = _AD_ListExchangeServers()
		If @error <> 0 Then Return SetError(3, @error, 0)
		$sAD_EMailServer = $aAD_Temp[1]
	EndIf
	Local $sAD_Mailboxpath = "LDAP://CN="
	$sAD_Mailboxpath &= $sAD_StoreName
	$sAD_Mailboxpath &= ",CN="
	$sAD_Mailboxpath &= $sAD_Store
	$sAD_Mailboxpath &= ",CN=InformationStore"
	$sAD_Mailboxpath &= ",CN="
	$sAD_Mailboxpath &= $sAD_EMailServer
	$sAD_Mailboxpath &= ",CN=Servers,CN="
	$sAD_Mailboxpath &= $sAD_AdminGroup
	$sAD_Mailboxpath &= ",CN=Administrative Groups,CN=" & $sAD_EmailDomain & ",CN=Microsoft Exchange,CN=Services,CN=Configuration,"
	$sAD_Mailboxpath &= $sAD_DNSDomain
	$oAD_User.CreateMailbox($sAD_Mailboxpath)
	$oAD_User.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_CreateMailbox

; #FUNCTION#====================================================================================================================
; Name...........: _AD_DeleteMailbox
; Description ...: Deletes a users mailbox.
; Syntax.........: _AD_DeleteMailbox($sAD_User)
; Parameters ....: $sAD_User - User account (sAMAccountName or FQDN) to delete mailbox from
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_User does not exist
;                  |2 - $sAD_User doesn't have a mailbox
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_DeleteMailbox($sAD_User)

	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_User, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_User & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_User = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	If $oAD_User.HomeMDB = "" Then Return SetError(2, 0, 0)
	$oAD_User.DeleteMailbox
	$oAD_User.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_DeleteMailbox

; #FUNCTION#====================================================================================================================
; Name...........: _AD_MailEnableGroup
; Description ...: Mail-enables a Group.
; Syntax.........: _AD_MailEnableGroup($sAD_Group)
; Parameters ....: $sAD_Group - Groupname (sAMAccountName or FQDN)
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - $sAD_Group does not exist
;                  |x - Error returned by SetInfo function (Missing permission etc.)
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_MailEnableGroup($sAD_Group)

	If Not _AD_ObjectExists($sAD_Group) Then Return SetError(1, 0, 0)
	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Group, 3, 1) = "=" Then $sAD_Property = "distinguishedName"; FQDN provided
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Group & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Group = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	$oAD_Group.MailEnable
	$oAD_Group.Put("grouptype", $ADS_GROUP_TYPE_UNIVERSAL_SECURITY)
	$oAD_Group.SetInfo
	If @error <> 0 Then Return SetError(@error, 0, 0)
	Return 1

EndFunc   ;==>_AD_MailEnableGroup

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_IsAccountExpired
; Description ...: Returns 1 if the account (user, computer) has expired.
; Syntax.........: _AD_IsAccountExpired([$sAD_Object = @Username])
; Parameters ....: $sAD_Object - Optional: Account (User, computer) to check (default = @Username). Can be specified as Fully Qualified Domain Name (FQDN) or sAMAccountName
; Return values .: Success - 1, The specified account has expired
;                  Failure - 0, sets @error to:
;                  |0 - Account has not expired
;                  |1 - $sAD_Object could not be found
;                  |2 - An error occurred when accessing property AccountExpirationDate. @extended is set to the error returned by LDAP
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_GetAccountsExpired, _AD_SetAccountExpire
; Link ..........: http://www.rlmueller.net/AccountExpires.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_IsAccountExpired($sAD_Object = @UserName)

	Local $sAD_Property = "sAMAccountName"
	If StringMid($sAD_Object, 3, 1) = "=" Then $sAD_Property = "distinguishedName" ; FQDN provided
	If _AD_ObjectExists($sAD_Object, $sAD_Property) = 0 Then Return SetError(1, 0, 0)
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	Local $sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
	Local $oAD_Object = __AD_ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
	Local $sAD_AccountExpires = $oAD_Object.AccountExpirationDate
	If @error Then Return SetError(2, @error, 0)
	If $sAD_AccountExpires <> "19700101000000" And StringLeft($sAD_AccountExpires, 8) <> "16010101" And $sAD_AccountExpires < @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC Then Return 1
	Return

EndFunc   ;==>_AD_IsAccountExpired

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetAccountsExpired
; Description ...: Returns an array with FQDNs of expired accounts (user, computer).
; Syntax.........: _AD_GetAccountsExpired([$sAD_Class = "user"[ ,$sAD_DTEExpire = ""[ ,$sAD_DTSExpire = ""[, $sAD_Root = ""]]]])
; Parameters ....: $sAD_Class - Optional: Specifies if expired user accounts or computer accounts should be returned (default = "user").
;                  |"user" - Returns objects of category "user"
;                  |"computer" - Returns objects of category "computer"
;                  $sAD_DTEExpire - YYYY/MM/DD HH:MM:SS (local time) returns all accounts that expire between $sAD_DTSExpire and the specified date/time (default = "" = Now)
;                  $sAD_DTSExpire - YYYY/MM/DD HH:MM:SS (local time) returns all accounts that expire between the specified date/time and $sAD_DTEExpire (default = "1601/01/01 00:00:00)
;                  $sAD_Root - Optional: FQDN of the OU where the search should start (default = "" = search the whole tree)
; Return values .: Success - One-based two dimensional array of FQDNs of expired accounts
;                  |0 - FQDNs of expired accounts
;                  |1 - account expired YYYY/MM/DD HH:NMM:SS UTC
;                  |2 - account expired YYYY/MM/DD HH:NMM:SS local time of calling user
;                  Failure - "", sets @error to:
;                  |1 - No expired accounts found. @extended is set to the error returned by LDAP
;                  |2 - Specified date/time is invalid
;                  |3 - Invalid value for $sAD_Class. Has to be "user" or "computer"
;                  |4 - Specified $sAD_Root does not exist
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _AD_IsAccountExpired, _AD_SetAccountExpire
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetAccountsExpired($sAD_Class = "user", $sAD_DTEExpire = "", $sAD_DTSExpire = "", $sAD_Root = "")

	If $sAD_Class <> "user" And $sAD_Class <> "computer" Then Return SetError(3, 0, 0)
	If $sAD_Root = "" Then
		$sAD_Root = $sAD_DNSDomain
	Else
		If _AD_ObjectExists($sAD_Root, "distinguishedName") = 0 Then Return SetError(4, 0, "")
	EndIf
	; process end date/time
	If $sAD_DTEExpire = "" Then
		$sAD_DTEExpire = _Date_Time_GetSystemTime() ; Get current date/time (UTC)
		$sAD_DTEExpire = _Date_Time_SystemTimeToDateTimeStr($sAD_DTEExpire, 1) ; convert to format yyyy/mm/dd hh:mm:ss
	ElseIf Not _DateIsValid($sAD_DTEExpire) Then
		Return SetError(2, 0, 0)
	Else
		$sAD_DTEExpire = _Date_Time_EncodeSystemTime(StringMid($sAD_DTEExpire, 6, 2), StringMid($sAD_DTEExpire, 9, 2), StringLeft($sAD_DTEExpire, 4), _ ; encode input
				StringMid($sAD_DTEExpire, 12, 2), StringMid($sAD_DTEExpire, 15, 2), StringMid($sAD_DTEExpire, 18, 2))
		Local $sAD_DTEExpireUTC = _Date_Time_TzSpecificLocalTimeToSystemTime(DllStructGetPtr($sAD_DTEExpire)) ; convert local time to UTC
		$sAD_DTEExpire = _Date_Time_SystemTimeToDateTimeStr($sAD_DTEExpireUTC, 1) ; convert to format yyyy/mm/dd hh:mm:ss
	EndIf
	; process start date/time
	If $sAD_DTSExpire = "" Then $sAD_DTSExpire = "1600/01/01 00:00:00"
	If Not _DateIsValid($sAD_DTSExpire) Then
		Return SetError(2, 0, 0)
	Else
		$sAD_DTSExpire = _Date_Time_EncodeSystemTime(StringMid($sAD_DTSExpire, 6, 2), StringMid($sAD_DTSExpire, 9, 2), StringLeft($sAD_DTSExpire, 4), _ ; encode input
				StringMid($sAD_DTSExpire, 12, 2), StringMid($sAD_DTSExpire, 15, 2), StringMid($sAD_DTSExpire, 18, 2))
		Local $sAD_DTSExpireUTC = _Date_Time_TzSpecificLocalTimeToSystemTime(DllStructGetPtr($sAD_DTSExpire)) ; convert local time to UTC
		$sAD_DTSExpire = _Date_Time_SystemTimeToDateTimeStr($sAD_DTSExpireUTC, 1) ; convert to format yyyy/mm/dd hh:mm:ss
	EndIf
	Local $iAD_DTEExpire = _DateDiff("s", "1601/01/01 00:00:00", $sAD_DTEExpire) * 10000000 ; convert end date/time to Integer8
	Local $iAD_DTSExpire = _DateDiff("s", "1601/01/01 00:00:00", $sAD_DTSExpire) * 10000000 ; convert start date/time to Integer8
	Local $iAD_Temp, $sAD_Temp
	Local $sAD_DTStruct = DllStructCreate("dword low;dword high")
	; -1 to remove rounding errors
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_Root & ">;(&(objectCategory=person)(objectClass=" & $sAD_Class & ")" & _
			"(!accountExpires=0)(accountExpires<=" & Int($iAD_DTEExpire) - 1 & ")(accountExpires>=" & Int($iAD_DTSExpire) - 1 & "));distinguishedName,accountExpires;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute
	If @error Or Not IsObj($oAD_RecordSet) Or $oAD_RecordSet.RecordCount = 0 Then Return SetError(1, @error, "")
	Local $aAD_FQDN[$oAD_RecordSet.RecordCount + 1][3]
	$aAD_FQDN[0][0] = $oAD_RecordSet.RecordCount
	Local $iAD_Count = 1
	While Not $oAD_RecordSet.EOF
		$aAD_FQDN[$iAD_Count][0] = $oAD_RecordSet.Fields(0).Value ; distinguishedName
		$iAD_Temp = $oAD_RecordSet.Fields(1).Value ; accountExpires
		DllStructSetData($sAD_DTStruct, "Low", $iAD_Temp.LowPart)
		DllStructSetData($sAD_DTStruct, "High", $iAD_Temp.HighPart)
		$sAD_Temp = _Date_Time_FileTimeToSystemTime(DllStructGetPtr($sAD_DTStruct))
		$aAD_FQDN[$iAD_Count][1] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp, 1) ; accountExpires as UTC
		$sAD_Temp = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sAD_Temp))
		$aAD_FQDN[$iAD_Count][2] = _Date_Time_SystemTimeToDateTimeStr($sAD_Temp, 1) ; accountExpires as local time
		$iAD_Count += 1
		$oAD_RecordSet.MoveNext
	WEnd
	$aAD_FQDN[0][0] = UBound($aAD_FQDN) - 1
	Return $aAD_FQDN

EndFunc   ;==>_AD_GetAccountsExpired

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ListSchemaVersions
; Description ...: Returns information about the AD Schema versions.
; Syntax.........: _AD_ListSchemaVersions()
; Parameters ....: None
; Return values .: Success - Returns an one-based one dimensional array of the following Schema versions.
;                  |1 - Active Directory Schema version. This can be one of the following values:
;                       13 - Windows 2000 Server
;                       30 - Windows Server 2003 RTM / Service Pack 1 / Service Pack 2
;                       31 - Windows Server 2003 R2
;                       44 - Windows Server 2008 RTM
;                       47 - Windows Server 2008 R2
;                  |2 - Exchange Schema version. This can be one of the following values:
;                       4397 - Exchange Server 2000 RTM
;                       4406 - Exchange Server 2000 With Service Pack 3
;                       6870 - Exchange Server 2003 RTM
;                       6936 - Exchange Server 2003 With Service Pack 3
;                       10628 - Exchange Server 2007
;                       11116 - Exchange 2007 With Service Pack 1
;                       14622 - Exchange 2007 With Service Pack 2, Exchange 2010 RTM
;                       14625 - Exchange 2007 SP3
;                       14720 - Exchange 2010 SP1 (beta)
;                       14726 - Exchange 2010 SP1
;                  |3 - Exchange 2000 Active Directory Connector version. This can be one of the following values:
;                       4197 - Exchange Server 2000 RTM
;                  |4 - Office Communications Server Schema version. This can be one of the following values:
;                       1006 - LCS 2005 SP1
;                       1007 - OCS 2007
;                       1008 - OCS 2007 R2
;                       1100 - Lync Server 2010
; Author ........: water
; Modified.......:
; Remarks .......: RTM stands for "Release to Manufacturing"
; Related .......:
; Link ..........: http://blog.dikmenoglu.de/Die+AD+Und+Exchange+Schemaversion+Abfragen.aspx, http://www.msxfaq.de/admin/build.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ListSchemaVersions()

	Local $aAD_Result[5] = [4]
	Local $sAD_LDAPEntry
	; Active Directory Schema Version
	Local $sAD_SchemaNamingContext = $__oAD_RootDSE.Get("SchemaNamingContext")
	Local $oAD_Object = __AD_ObjGet("LDAP://" & $sAD_SchemaNamingContext) ; Retrieve the COM Object for the object
	$aAD_Result[1] = $oAD_Object.Get("objectVersion")
	; Exchange Schema version
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_SchemaNamingContext & ">;(name=ms-Exch-Schema-Version-Pt);ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	If IsObj($oAD_RecordSet) And $oAD_RecordSet.RecordCount > 0 Then
		$sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
		$oAD_Object = ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
		$aAD_Result[2] = $oAD_Object.Get("rangeUpper")
	EndIf
	; Exchange 2000 Active Directory Connector version
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_SchemaNamingContext & ">;(name=ms-Exch-Schema-Version-Adc);ADsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	If IsObj($oAD_RecordSet) And $oAD_RecordSet.RecordCount > 0 Then
		$sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
		$oAD_Object = ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
		$aAD_Result[3] = $oAD_Object.Get("rangeUpper")
	EndIf
	; Office Communications Server Schema version
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_SchemaNamingContext & ">;(name=ms-RTC-SIP-SchemaVersion);ADsPath;subtree"
	$oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
	If IsObj($oAD_RecordSet) And $oAD_RecordSet.RecordCount > 0 Then
		$sAD_LDAPEntry = $oAD_RecordSet.fields(0).value
		$oAD_Object = ObjGet($sAD_LDAPEntry) ; Retrieve the COM Object for the object
		$aAD_Result[4] = $oAD_Object.Get("rangeUpper")
	EndIf

	Return $aAD_Result

EndFunc   ;==>_AD_ListSchemaVersions

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_ObjectExistsInSchema
; Description ...: Returns 1 if exactly one object exists for the given property in the Active Directory Schema.
; Syntax.........: _AD_ObjectExistsInSchema($sAD_Object [, $sAD_Property = "LDAPDisplayName"])
; Parameters ....: $sAD_Object   - Optional: Object to check
;                  $sAD_Property - Optional: Property to check (default = LDAPDisplayName)
; Return values .: Success - 1, Exactly one object exists for the given property in the Active Directory Schema
;                  Failure - 0, sets @error to:
;                  |1 - No object found for the specified property
;                  |x - More than one object found for the specified property. x is the number of objects found
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_ObjectExistsInSchema($sAD_Object, $sAD_Property = "LDAPDisplayName")

	Local $sAD_SchemaNamingContext = $__oAD_RootDSE.Get("SchemaNamingContext")
	$__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_SchemaNamingContext & ">;(" & $sAD_Property & "=" & $sAD_Object & ");ADsPath;subtree"
	Local $oAD_RecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object, if it exists
	If IsObj($oAD_RecordSet) Then
		If $oAD_RecordSet.RecordCount = 1 Then
			Return 1
		ElseIf $oAD_RecordSet.RecordCount > 1 Then
			Return SetError($oAD_RecordSet.RecordCount, 0, 0)
		Else
			Return SetError(1, 0, 0)
		EndIf
	Else
		Return SetError(1, 0, 0)
	EndIf

EndFunc   ;==>_AD_ObjectExistsInSchema

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_FixSpecialChars
; Description ...: Returns either corrected (with 'escaped' chars) or uncorrected (removes 'escapes' for chars) text.
; Syntax.........: _AD_FixSpecialChars($sAD_Text[, $iAD_Option = 0[, $sAD_EscapeChar = '"\/#,+<>;=']])
; Parameters ....: $sAD_Text - Text to add / remove escape characters
;                  $iAD_Option - Optional: Defines whether to insert the escape character (Default) or remove them.
;                                0 = Insert the escape characters (default)
;                                1 = Remove the escape characters
;                  $sAD_EscapeChar - Optional: List of characters to escape or unescape (default = '"\/#,+<>;=')
; Return values .: $sAD_Text with inserted or removed escape characters
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: Leading or trailing spaces have to be escaped manually by the user
; Related .......:
; Link ..........: http://www.rlmueller.net/CharactersEscaped.htm
; Example .......: Yes
; ===============================================================================================================================
Func _AD_FixSpecialChars($sAD_Text, $iAD_Option = 0, $sAD_EscapeChar = '"\/#,+<>;=')

	If $iAD_Option = 0 Then
		; "?<!" means: Lookbehind assertion (negative) - see: http://www.autoitscript.com/forum/index.php?showtopic=112674&view=findpost&p=789310
		$sAD_Text = StringRegExpReplace($sAD_Text, "(?<!\\)([" & $sAD_EscapeChar & "])", "\\$1")
	Else
		If StringInStr($sAD_EscapeChar, '"') Then $sAD_Text = StringReplace($sAD_Text, '\"', '"')
		If StringInStr($sAD_EscapeChar, '/') Then $sAD_Text = StringReplace($sAD_Text, '\/', '/')
		If StringInStr($sAD_EscapeChar, '#') Then $sAD_Text = StringReplace($sAD_Text, '\#', '#')
		If StringInStr($sAD_EscapeChar, ',') Then $sAD_Text = StringReplace($sAD_Text, '\,', ',')
		If StringInStr($sAD_EscapeChar, '+') Then $sAD_Text = StringReplace($sAD_Text, '\+', '+')
		If StringInStr($sAD_EscapeChar, '<') Then $sAD_Text = StringReplace($sAD_Text, '\<', '<')
		If StringInStr($sAD_EscapeChar, '>') Then $sAD_Text = StringReplace($sAD_Text, '\>', '>')
		If StringInStr($sAD_EscapeChar, ';') Then $sAD_Text = StringReplace($sAD_Text, '\;', ';')
		If StringInStr($sAD_EscapeChar, '=') Then $sAD_Text = StringReplace($sAD_Text, '\=', '=')
		If StringInStr($sAD_EscapeChar, '\') Then $sAD_Text = StringReplace($sAD_Text, '\\', '\')
	EndIf
	Return $sAD_Text

EndFunc   ;==>_AD_FixSpecialChars

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetLastADSIError
; Description ...: Retrieve the calling thread's last ADSI error code value.
; Syntax.........: _AD_GetLastADSIError()
; Parameters ....: None
; Return values .: Success - A one-based array containing the following values:
;                  |1 - ADSI error code (decimal)
;                  |2 - Unicode string that describes the error
;                  |3 - name of the provider that raised the error
;                  |4 - Win32 error code extracted from element[2]
;                  |5 - description of the Win32 error code as returned by _WinAPI_FormatMessage
;                  Failure - "", sets @error to the return value of DLLCall
; Author ........: water, card0384
; Modified.......:
; Remarks .......: This and more errors will be handled (error codes are in hex):
;                  525 - user not found
;                  52e - invalid credentials
;                  530 - not permitted to logon at this time
;                  532 - password expired
;                  533 - account disabled
;                  701 - account expired
;                  773 - user must reset password
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/cc231199(PROT.10).aspx (Win32 Error Codes), http://forums.sun.com/thread.jspa?threadID=703398
; Example .......:
; ===============================================================================================================================
Func _AD_GetLastADSIError()

	Local $aAD_LastError[6] = [5]
	Local $EC = DllStructCreate("DWord")
	Local $ED = DllStructCreate("wchar[256]")
	Local $PN = DllStructCreate("wchar[256]")
	; ADsGetLastError: http://msdn.microsoft.com/en-us/library/aa772183(VS.85).aspx
	DllCall("Activeds.dll", "DWORD", "ADsGetLastError", "ptr", DllStructGetPtr($EC), "ptr", DllStructGetPtr($ED), "DWORD", 256, "ptr", DllStructGetPtr($PN), "DWORD", 256)
	If @error <> 0 Then Return SetError(@error, @extended, "")
	$aAD_LastError[1] = DllStructGetData($EC, 1) ; error code (decimal)
	$aAD_LastError[2] = DllStructGetData($ED, 1) ; Unicode string that describes the error
	$aAD_LastError[3] = DllStructGetData($PN, 1) ; name of the provider that raised the error
	; Old version to set element 4
	;	Local $sAD_Error = StringTrimLeft($aAD_LastError[2], StringInStr($aAD_LastError[2], "AcceptSecurityContext", 2))
	;	$sAD_Error = StringTrimLeft($sAD_Error, StringInStr($sAD_Error, " data", 2) + 5)
	;	$aAD_LastError[4] = StringTrimRight($sAD_Error, StringLen($sAD_Error) - StringInStr($sAD_Error, ", vece", 2) + 1)
	If $aAD_LastError[2] <> "" Then
		Local $aAD_TempError = StringSplit($aAD_LastError[2], ",")
		$aAD_TempError = StringSplit(StringStripWS($aAD_TempError[3], 7), " ")
		$aAD_LastError[4] = $aAD_TempError[2]
	EndIf
	_WinAPI_FormatMessage($__WINAPICONSTANT_FORMAT_MESSAGE_FROM_SYSTEM, 0, Dec($aAD_LastError[4]), 0, $aAD_LastError[5], 4096, 0)
	Return $aAD_LastError

EndFunc   ;==>_AD_GetLastADSIError

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetADOProperties
; Description ...: Retrieves all properties of an ADO object (Connection, Command).
; Syntax.........: _AD_GetADOProperties($sAD_ADOobject[, $sAD_Properties = ""])
; Parameters ....: $sAD_ADOobject - ADO object for which to retrieve the properties. Can be either "Connection" or "Command"
;                  $sAD_Properties - Optional: Comma separated list of properties to return (default = "" = return all properties)
; Return values .: Success - Returns a one based two-dimensional array with all properties and their values of the specified object
;                  Failure - "", sets @error to:
;                  |1 - $sAD_ADOObject is invalid. Should either be "Command" or "Connection"
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........: ADO Command: http://msdn.microsoft.com/en-us/library/ms675022(v=VS.85).aspx, ADO Connection: http://msdn.microsoft.com/en-us/library/ms681546(v=VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetADOProperties($sAD_ADOObject, $sAD_Properties = "")

	Local $oAD_Object
	If StringLeft($sAD_ADOObject, 3) = "Com" Then
		$oAD_Object = $__oAD_Command
	ElseIf StringLeft($sAD_ADOObject, 3) = "Con" Then
		$oAD_Object = $__oAD_Connection
	Else
		Return SetError(1, 0, "")
	EndIf
	$sAD_Properties = "," & $sAD_Properties & ","
	Local $aAD_Properties[$oAD_Object.Properties.Count + 1][2] = [[$oAD_Object.Properties.Count, 2]], $iAD_Index = 1
	For $oAD_Property In $oAD_Object.Properties
		If Not ($sAD_Properties = ",," Or StringInStr($sAD_Properties, "," & $oAD_Property.Name & ",") > 0) Then ContinueLoop
		$aAD_Properties[$iAD_Index][0] = $oAD_Property.Name
		$aAD_Properties[$iAD_Index][1] = $oAD_Property.Value
		$iAD_Index += 1
	Next
	ReDim $aAD_Properties[$iAD_Index][2]
	$aAD_Properties[0][0] = $iAD_Index - 1
	Return $aAD_Properties

EndFunc   ;==>_AD_GetADOProperties

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_SetADOProperties
; Description ...: Sets properties of an ADO command object.
; Syntax.........: _AD_SetADOProperties($sAD_P1 = ""[, $sAD_P2 = ""[, $sAD_P3 = ""[, $sAD_P4 = ""[, $sAD_P5 = ""[, $sAD_P6 = ""[, $sAD_P7 = ""[, $sAD_P8 = ""[, $sAD_P9 = ""[, $sAD_P10 = ""]]]]]]]]])
; Parameters ....: $sAD_P1        - Property to set. This can be a string with the format propertyname=value OR
;                  +a zero based one-dimensional array with unlimited number of strings in the format propertyname=value
;                  $sAD_P2        - Optional: Same as $sAD_P1 but no array is allowed
;                  $sAD_P3        - Optional: Same as $sAD_P2
;                  $sAD_P4        - Optional: Same as $sAD_P2
;                  $sAD_P5        - Optional: Same as $sAD_P2
;                  $sAD_P6        - Optional: Same as $sAD_P2
;                  $sAD_P7        - Optional: Same as $sAD_P2
;                  $sAD_P8        - Optional: Same as $sAD_P2
;                  $sAD_P9        - Optional: Same as $sAD_P2
;                  $sAD_P10       - Optional: Same as $sAD_P2
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |1 - Invalid format of the parameter. Has to be propertyname=value. @extended = number of the invalid parameter (zero based)
;                  |x - Error setting the specified property. @extended = number of the invalid parameter (zero based)
; Author ........: water
; Modified.......:
; Remarks .......: You can query the properties of the ADO connection and command object but you can only set the properties of the command object.
;                  +After the connection is opened by _AD_Open the properties of the connection object are read only.
; Related .......:
; Link ..........: ADO Command: http://msdn.microsoft.com/en-us/library/ms675022(v=VS.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _AD_SetADOProperties($sAD_P1 = "", $sAD_P2 = "", $sAD_P3 = "", $sAD_P4 = "", $sAD_P5 = "", $sAD_P6 = "", $sAD_P7 = "", $sAD_P8 = "", $sAD_P9 = "", $sAD_P10 = "")

	Local $aAD_Properties[10]
	; Move properties into an array
	If Not IsArray($sAD_P1) Then
		$aAD_Properties[0] = $sAD_P1
		$aAD_Properties[1] = $sAD_P2
		$aAD_Properties[2] = $sAD_P3
		$aAD_Properties[3] = $sAD_P4
		$aAD_Properties[4] = $sAD_P5
		$aAD_Properties[5] = $sAD_P6
		$aAD_Properties[6] = $sAD_P7
		$aAD_Properties[7] = $sAD_P8
		$aAD_Properties[8] = $sAD_P9
		$aAD_Properties[9] = $sAD_P10
	Else
		$aAD_Properties = $sAD_P1
	EndIf
	; set properties
	For $iAD_Index = 0 To UBound($aAD_Properties) - 1
		If $aAD_Properties[$iAD_Index] = "" Then ContinueLoop
		Local $aAD_Temp = StringSplit($aAD_Properties[$iAD_Index], "=")
		If @error = 1 Then Return SetError(1, $iAD_Index, 0)
		$aAD_Temp[1] = StringStripWS($aAD_Temp[1], 7)
		$__oAD_Command.Properties($aAD_Temp[1]) = $aAD_Temp[2]
		If @error <> 0 Then Return SetError(@error, $iAD_Index, 0)
	Next
	Return 1

EndFunc   ;==>_AD_SetADOProperties

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_VersionInfo
; Description ...: Returns an array of information about the AD.au3 UDF.
; Syntax.........: _AD_VersionInfo()
; Parameters ....: None
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Release Type (T=Test or V=Production)
;                  |2 - Major Version
;                  |3 - Minor Version
;                  |4 - Sub Version
;                  |5 - Release Date (YYYYMMDD)
;                  |6 - AutoIt version required
;                  |7 - List of authors separated by ","
;                  |8 - List of contributors separated by ","
; Author ........: water
; Modified.......:
; Remarks .......: Based on function _IE_VersionInfo written bei Dale Hohm
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_VersionInfo()

	Local $aOL_VersionInfo[9] = [8, "V", 1, 3, 0.0, "20121012", "3.3.6.0", "Jonathan Clelland, water", _
			"feeks, KenE, Sundance, supersonic, Talder, Joe2010, Suba, Ethan Turk, Jerold Schulman, Stephane, card0384"]
	Return $aOL_VersionInfo

EndFunc   ;==>_AD_VersionInfo

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_Int8ToSec
; Description ...: Function to convert Integer8 attributes from 64-bit numbers to seconds.
; Syntax.........: __AD_Int8ToSec($iAD_Low, $iAD_High)
; Parameters ....: $iAD_Low - Lower Part of the Large Integer
;                  $iAD_High - Higher Part of the Large Integer
; Return values .: Integer (Double Word) value
; Author ........: Jerold Schulman
; Modified.......: water
; Remarks .......: This function is used internally
;                  Many attributes in Active Directory have a data type (syntax) called Integer8.
;                  These 64-bit numbers (8 bytes) often represent time in 100-nanosecond intervals. If the Integer8 attribute is a date,
;                  the value represents the number of 100-nanosecond intervals since 12:00 AM January 1, 1601.
; Related .......:
; Link ..........: http://www.rlmueller.net/Integer8Attributes.htm, http://msdn.microsoft.com/en-us/library/cc208659.aspx
; Example .......:
; ===============================================================================================================================
Func __AD_Int8ToSec($oAD_Int8)

	Local $lngHigh, $lngLow
	$lngHigh = $oAD_Int8.HighPart
	$lngLow = $oAD_Int8.LowPart
	If $lngLow < 0 Then
		$lngHigh = $lngHigh + 1
	EndIf
	Return -($lngHigh * (2 ^ 32) + $lngLow) / (10000000)

EndFunc   ;==>__AD_Int8ToSec

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_LargeInt2Double
; Description ...: Converts a large Integer value to an Integer (Double Word) value.
; Syntax.........: __AD_LargeInt2Double($iAD_Low, $iAD_High)
; Parameters ....: $iAD_Low - Lower Part of the Large Integer
;                  $iAD_High - Higher Part of the Large Integer
; Return values .: Integer (Double Word) value
; Author ........: Sundance
; Modified.......: water
; Remarks .......: This function is used internally
; Related .......:
; Link ..........: http://www.autoitscript.com/forum/index.php?showtopic=49627&view=findpost&p=422402
; Example .......:
; ===============================================================================================================================
Func __AD_LargeInt2Double($iAD_Low, $iAD_High)

	Local $iAD_ResultLow, $iAD_ResultHigh
	If $iAD_Low < 0 Then
		$iAD_ResultLow = 2 ^ 32 + $iAD_Low
	Else
		$iAD_ResultLow = $iAD_Low
	EndIf
	If $iAD_High < 0 Then
		$iAD_ResultHigh = 2 ^ 32 + $iAD_High
	Else
		$iAD_ResultHigh = $iAD_High
	EndIf
	Return $iAD_ResultLow + $iAD_ResultHigh * 2 ^ 32

EndFunc   ;==>__AD_LargeInt2Double

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_ObjGet
; Description ...: Returns an LDAP object from a FQDN.
;                  Will use the alternative credentials $sAD_UserId/$sAD_Password if they are set.
; Syntax.........: __AD_ObjGet($sAD_FQDN)
; Parameters ....: $sAD_FQDN - Fully Qualified Domain name of the object for which the LDAP object will be returned.
; Return values .: LDAP object
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This function is used internally
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772247(VS.85).aspx (ADS_AUTHENTICATION_ENUM Enumeration)
; Example .......:
; ===============================================================================================================================
Func __AD_ObjGet($sAD_FQDN)

	If $sAD_UserId = "" Then
		Return ObjGet($sAD_FQDN)
	Else
		Return $__oAD_OpenDS.OpenDSObject($sAD_FQDN, $sAD_UserId, $sAD_Password, $__bAD_BindFlags)
	EndIf

EndFunc   ;==>__AD_ObjGet

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_ErrorHandler
; Description ...: Internal Error event handler for COM errors.
; Syntax.........: __AD_ErrorHandler()
; Parameters ....: None
; Return values .: @error is set to the COM error by AutoIt
; Author ........: Jonathan Clelland
; Modified.......: water
; Remarks .......: This function is used internally
;                  0x80005xxx - ADSI error codes
;                  0x80007xxx - LDAP error codes for ADSI
;
;                  ADSI Error Codes:
;                  Code        Symbol                        Description
;                  -----------------------------------------------------
;                  0x80005000  E_ADS_BAD_PATHNAME            An invalid ADSI path name was passed
;                  0x80005001  E_ADS_INVALID_DOMAIN_OBJECT   An unknown ADSI domain object was requested
;                  0x80005002  E_ADS_INVALID_USER_OBJECT     An unknown ADSI user object was requested
;                  0x80005003  E_ADS_INVALID_COMPUTER_OBJECT An unknown ADSI computer object was requested
;                  0x80005004  E_ADS_UNKNOWN_OBJECT          An unknown ADSI object was requested
;                  0x80005005  E_ADS_PROPERTY_NOT_SET        The specified ADSI property was not set
;                  0x80005006  E_ADS_PROPERTY_NOT_SUPPORTED  The specified ADSI property is not supported
;                  0x80005007  E_ADS_PROPERTY_INVALID        The specified ADSI property is invalid
;                  0x80005008  E_ADS_BAD_PARAMETER           One or more input parameters are invalid
;                  0x80005009  E_ADS_OBJECT_UNBOUND          The specified ADSI object is not bound to a remote resource
;                  0x8000500A  E_ADS_PROPERTY_NOT_MODIFIED   The specified ADSI object has not been modified
;                  0x8000500B  E_ADS_PROPERTY_MODIFIED       The specified ADSI object has been modified
;                  0x8000500C  E_ADS_CANT_CONVERT_DATATYPE   The ADSI data type cannot be converted to/from a native DS data type
;                  0x8000500D  E_ADS_PROPERTY_NOT_FOUND      The ADSI property cannot be found in the cache
;                  0x8000500E  E_ADS_OBJECT_EXISTS           The ADSI object exists
;                  0x8000500F  E_ADS_SCHEMA_VIOLATION        The attempted action violates the directory service schema rules
;                  0x80005010  E_ADS_COLUMN_NOT_SET          The specified column in the ADSI was not set
;                  0x00005011  S_ADS_ERRORSOCCURRED          One or more errors occurred
;                  0x00005012  S_ADS_NOMORE_ROWS             The search operation has reached the last row
;                  0x00005013  S_ADS_NOMORE_COLUMNS          The search operation has reached the last column for the current row
;                  0x80005014  E_ADS_INVALID_FILTER          The specified search filter is invalid
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/aa772195(VS.85).aspx (Active Directory Services Interface Error Codes)
;                  http://msdn.microsoft.com/en-us/library/cc704587(PROT.10).aspx (HRESULT Values)
; Example .......:
; ===============================================================================================================================
Func __AD_ErrorHandler()

	Local $bAD_HexNumber = Hex($__oAD_MyError.number, 8)
	Local $aAD_VersionInfo = _AD_VersionInfo()
	Local $sAD_Error = "COM Error Encountered in " & @ScriptName & @CRLF & _
			"AD UDF version = " & $aAD_VersionInfo[2] & "." & $aAD_VersionInfo[3] & "." & $aAD_VersionInfo[4] & @CRLF & _
			"@AutoItVersion = " & @AutoItVersion & @CRLF & _
			"@AutoItX64 = " & @AutoItX64 & @CRLF & _
			"@Compiled = " & @Compiled & @CRLF & _
			"@OSArch = " & @OSArch & @CRLF & _
			"@OSVersion = " & @OSVersion & @CRLF & _
			"Scriptline = " & $__oAD_MyError.scriptline & @CRLF & _
			"NumberHex = " & $bAD_HexNumber & @CRLF & _
			"Number = " & $__oAD_MyError.number & @CRLF & _
			"WinDescription = " & StringStripWS($__oAD_MyError.WinDescription, 2) & @CRLF & _
			"Description = " & StringStripWS($__oAD_MyError.description, 2) & @CRLF & _
			"Source = " & $__oAD_MyError.Source & @CRLF & _
			"HelpFile = " & $__oAD_MyError.HelpFile & @CRLF & _
			"HelpContext = " & $__oAD_MyError.HelpContext & @CRLF & _
			"LastDllError = " & $__oAD_MyError.LastDllError
	If $__iAD_Debug > 0 Then
		If $__iAD_Debug = 1 Then ConsoleWrite($sAD_Error & @CRLF & "========================================================" & @CRLF)
		If $__iAD_Debug = 2 Then MsgBox(64, "Active Directory Functions - Debug Info", $sAD_Error)
		If $__iAD_Debug = 3 Then FileWrite($__sAD_DebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
				"-------------------" & @CRLF & $sAD_Error & @CRLF & "========================================================" & @CRLF)
	EndIf

EndFunc   ;==>__AD_ErrorHandler

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_ReorderACE
; Description ...: Reorders the ACEs in a DACL to meet MS recommandations.
; Syntax.........: __AD_ReorderACE($oAD_DACL)
; Parameters ....: $oAD_DACL - Discretionary Access Control List
; Return values .: A reordered DACL according to MS recommandation
; Author ........: water (based on VB code by Richard L. Mueller)
; Modified.......:
; Remarks .......: This function is used internally
;+
;                  The Active Directory Service Interfaces (ADSI) property cache on Microsoft Windows Server 2003 and on Microsoft Windows XP
;                  will correctly order the DACL before writing it back to the object. Reordering is only required on Microsoft Windows 2000.
;                  The proper order of ACEs in an ACL is as follows:
;                    Access-denied ACEs that apply to the object itself
;                    Access-denied ACEs that apply to a child of the object, such as a property set or property
;                    Access-allowed ACEs that apply to the object itself
;                    Access-allowed ACEs that apply to a child of the object, such as a property set or property
;                    All inherited ACEs
; Related .......:
; Link ..........: http://support.microsoft.com/kb/269159/en-us
; Example .......:
; ===============================================================================================================================
Func __AD_ReorderACE($oAD_DACL)

	; Reorder ACE's in DACL
	Local $oAD_NewDACL, $oAD_InheritedDACL, $oAD_AllowDACL, $oAD_DenyDACL, $oAD_AllowObjectDACL, $oAD_DenyObjectDACL, $oAD_ACE

	$oAD_NewDACL = ObjCreate("AccessControlList")
	$oAD_InheritedDACL = ObjCreate("AccessControlList")
	$oAD_AllowDACL = ObjCreate("AccessControlList")
	$oAD_DenyDACL = ObjCreate("AccessControlList")
	$oAD_AllowObjectDACL = ObjCreate("AccessControlList")
	$oAD_DenyObjectDACL = ObjCreate("AccessControlList")

	For $oAD_ACE In $oAD_DACL
		If ($oAD_ACE.AceFlags And $ADS_ACEFLAG_INHERITED_ACE) = $ADS_ACEFLAG_INHERITED_ACE Then
			$oAD_InheritedDACL.AddAce($oAD_ACE)
		Else
			Switch $oAD_ACE.AceType
				Case $ADS_ACETYPE_ACCESS_ALLOWED
					$oAD_AllowDACL.AddAce($oAD_ACE)
				Case $ADS_ACETYPE_ACCESS_DENIED
					$oAD_DenyDACL.AddAce($oAD_ACE)
				Case $ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
					$oAD_AllowObjectDACL.AddAce($oAD_ACE)
				Case $ADS_ACETYPE_ACCESS_DENIED_OBJECT
					$oAD_DenyObjectDACL.AddAce($oAD_ACE)
				Case Else
			EndSwitch
		EndIf
	Next
	For $oAD_ACE In $oAD_DenyDACL
		$oAD_NewDACL.AddAce($oAD_ACE)
	Next
	For $oAD_ACE In $oAD_DenyObjectDACL
		$oAD_NewDACL.AddAce($oAD_ACE)
	Next
	For $oAD_ACE In $oAD_AllowDACL
		$oAD_NewDACL.AddAce($oAD_ACE)
	Next
	For $oAD_ACE In $oAD_AllowObjectDACL
		$oAD_NewDACL.AddAce($oAD_ACE)
	Next
	For $oAD_ACE In $oAD_InheritedDACL
		$oAD_NewDACL.AddAce($oAD_ACE)
	Next
	$oAD_NewDACL.ACLRevision = $oAD_DACL.ACLRevision
	Return $oAD_NewDACL

EndFunc   ;==>__AD_ReorderACE
