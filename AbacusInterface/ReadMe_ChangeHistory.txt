Change History (Most Recent First)

Date		Author		Comment
=========== =========== ====================================================================== 
04/09/2019	DCB			Add CureatrAlert
01/07/2019	DCB			new process ReadPagerGCTransfer
09/06/18	DCB			Update ReadCapcodeGroup sub added new fields
01/26/17	DCB			Added PurgeOldCueLogFiles
08/18/14	DCB			CHANGED CREATE SSO USER logic. added new SSO_AM api key and mod logic to try and connect accounts even if the SSO
						SAO_AM account already exits.
						The authenticate process can set always use the SSO_IM api key regardless of the account being checking
						per Shawn.

07/28/14    DCB			MOD SSO API calls processes.  Check for valid JSON before deserialize, mods to error reporting to relay more information back
						, removed extra call to create user process
07/22/14	DCB			MOD changed myappUuid code in AddPrincipal process
07/08/14	DCB			MOD add logic to included Peter in change password error emails { Proxy errors}
07/02/14	DCB			MOD Added TryCatch to ChangePassword process when SSO Rest Call Fails
06/19/14	DCB			MOD SENDMail and ProcessErrormessage and send email
05/21/2014	DCB			Add HttpUtility.UrlEncode to all passwords for SSO API's
04/14/2014	DCB			Add Processing for SSO Acct Creation
11/18/13	DCB			Add ReadAppUserPasswordChange, ReadWIFCLFWD, ProcessErrorMessages, and support processes
10/30/13	DCB			Update to VS2010
						Moved to .Network Framework 4.0
						Added Newtonsoft.DLL, SSO.vb  {Need to write to SSO server in RestCall, ReadMe_ChangeHistory
						Added function APPUserPasswordChange
10/29/13	Larry		Many, many, many changes.

						