#!/bin/sh
#set -x

TOOL_NAME="Microsoft Office 365/2019/2016 Keychain Removal Tool"
TOOL_VERSION="2.7"

## Copyright (c) 2021 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary 
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com

# Constants
WORD2016PATH="/Applications/Microsoft Word.app"
EXCEL2016PATH="/Applications/Microsoft Excel.app"
POWERPOINT2016PATH="/Applications/Microsoft PowerPoint.app"
OUTLOOK2016PATH="/Applications/Microsoft Outlook.app"
ONENOTE2016PATH="/Applications/Microsoft OneNote.app"
SCRIPTPATH=$( cd $(dirname $0) ; pwd -P )

function ShowUsage {
# Shows tool usage and parameters
	echo $TOOL_NAME - $TOOL_VERSION
	echo "Purpose: Removes Office 365/2019/2016 for Mac keychain entries for 15.x and 16.x builds"
	echo "Usage: NukeOffKey [--Default] [--IRM] [--All] [--Force] [--Jamf] [--MAS]"
	echo "         [--Default] removes the logon, cache and ADAL keychain entries"
	echo "         [--IRM] removes the rights management keychain entries"
	echo "         [--All] removes logon, cache, ADAL, rights management and HelpShift keychain entries"
	echo "         [--Jamf] ignores the first 3 parameters for running via Jamf"
	echo "         [--MAS] removes the entries required for a clean CDN to Mac App Store conversion"
	echo
	exit 0
}

# Check that all keychain-integrated applications are not running
function CheckRunning {
	OPENAPPS=0
	WORDRUNSTATE=$(CheckLaunchState "$WORD2016PATH")
	if [ "$WORDRUNSTATE" == "1" ]; then
		OPENAPPS=$(($OPENAPPS + 1))
	fi
	EXCELRUNSTATE=$(CheckLaunchState "$EXCEL2016PATH")
	if [ "$EXCELRUNSTATE" == "1" ]; then
		OPENAPPS=$(($OPENAPPS + 1))
	fi
	POWERPOINTRUNSTATE=$(CheckLaunchState "$POWERPOINT2016PATH")
	if [ "$POWERPOINTRUNSTATE" == "1" ]; then
		OPENAPPS=$(($OPENAPPS + 1))
	fi
	OUTLOOKRUNSTATE=$(CheckLaunchState "$OUTLOOK2016PATH")
	if [ "$OUTLOOKRUNSTATE" == "1" ]; then
		OPENAPPS=$(($OPENAPPS + 1))
	fi
	ONENOTERUNSTATE=$(CheckLaunchState "$ONENOTE2016PATH")
	if [ "$ONENOTERUNSTATE" == "1" ]; then
		OPENAPPS=$(($OPENAPPS + 1))
	fi
	if [ "$OPENAPPS" != "0" ]; then
		echo "1"
	else
		echo "0"
	fi
}

# Checks to see if a process is running
function CheckLaunchState {
	APPPATH="$1"
	local RUNNING_RESULT=$(ps ax | grep -v grep | grep "$APPPATH")
	if [ "${#RUNNING_RESULT}" -gt 0 ]; then
		echo "1"
	else
		echo "0"
	fi
}

# Checks to see if the user has root-level permissions
function GetSudo {
	if [ "$EUID" != "0" ]; then
		sudo -p "Enter administrator password: " echo
		if [ $? -eq 0 ] ; then
			echo "0"
		else
			echo "1"
		fi
	fi
}

# Forcibly terminates a process
function ForceTerminate {
	PROCESS="$1"
	$(ps ax | grep -v grep | grep "$PROCESS" | awk '{print $1}' | xargs kill -9 2> /dev/null 1> /dev/null)
}

# Forcibly quits all Office 2019/2016 apps
function ForceQuit2016 {
	ForceTerminate "$WORD2016PATH"
	ForceTerminate "$EXCEL2016PATH"
	ForceTerminate "$POWERPOINT2016PATH"
	ForceTerminate "$OUTLOOK2016PATH"
	ForceTerminate "$ONENOTE2016PATH"
}

# Checks to see if 'Microsoft Office Credentials' for ADAL entries are present in the keychain
function FindEntryMsoCredentialSchemeADAL {
	/usr/bin/security find-internet-password -s 'msoCredentialSchemeADAL' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the first 'Microsoft Office Credentials' for ADAL entry from the keychain
function RemoveEntryMsoCredentialSchemeADAL {
	/usr/bin/security delete-internet-password -s 'msoCredentialSchemeADAL' 2> /dev/null 1> /dev/null
}

# Checks to see if 'Microsoft Office Credentials' for LiveID entries are present in the keychain
function FindEntryMsoCredentialSchemeLiveId {
	/usr/bin/security find-internet-password -s 'msoCredentialSchemeLiveId' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the first 'Microsoft Office Credentials' for LiveID entry from the keychain
function RemoveEntryMsoCredentialSchemeLiveId {
	/usr/bin/security delete-internet-password -s 'msoCredentialSchemeLiveId' 2> /dev/null 1> /dev/null
}

# Checks to see if 'Microsoft Office Credentials' for Unknown entries are present in the keychain
function FindEntryMsoCredentialSchemeUnknown {
	/usr/bin/security find-internet-password -s 'msoCredentialSchemeUnknown' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the first 'Microsoft Office Credentials' for Unknown entry from the keychain
function RemoveEntryMsoCredentialSchemeUnknown {
	/usr/bin/security delete-internet-password -s 'msoCredentialSchemeUnknown' 2> /dev/null 1> /dev/null
}

# Checks to see if 'MSOpenTech.ADAL.1*' entries are present in the keychain
function FindEntryMSOpenTechADAL1 {
	/usr/bin/security find-generic-password -G 'MSOpenTech.ADAL.1' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the first 'MSOpenTech.ADAL.1*' entry from the keychain
function RemoveEntryMSOpenTechADAL1 {
	/usr/bin/security delete-generic-password -G 'MSOpenTech.ADAL.1' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Identities Cache 2' entry is present in the keychain (15.x builds)
function FindEntryOfficeIdCache2 {
	/usr/bin/security find-generic-password -l 'Microsoft Office Identities Cache 2' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Identities Cache 2' entry from the keychain (15.x builds)
function RemoveEntryOfficeIdCache2 {
	/usr/bin/security delete-generic-password -l 'Microsoft Office Identities Cache 2' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Identities Cache 3' entry is present in the keychain (16.x builds)
function FindEntryOfficeIdCache3 {
	/usr/bin/security find-generic-password -l 'Microsoft Office Identities Cache 3' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Identities Cache 3' entry from the keychain (16.x builds)
function RemoveEntryOfficeIdCache3 {
	/usr/bin/security delete-generic-password -l 'Microsoft Office Identities Cache 3' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Identities Settings 2' entry is present in the keychain (15.x builds)
function FindEntryOfficeIdSettings2 {
	/usr/bin/security find-generic-password -l 'Microsoft Office Identities Settings 2' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Identities Settings 2' entry from the keychain (15.x builds)
function RemoveEntryOfficeIdSettings2 {
	/usr/bin/security delete-generic-password -l 'Microsoft Office Identities Settings 2' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Identities Settings 3' entry is present in the keychain (16.x builds)
function FindEntryOfficeIdSettings3 {
	/usr/bin/security find-generic-password -l 'Microsoft Office Identities Settings 3' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Identities Settings 3' entry from the keychain (16.x builds)
function RemoveEntryOfficeIdSettings3 {
	/usr/bin/security delete-generic-password -l 'Microsoft Office Identities Settings 3' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Ticket Cache' entry is present in the keychain (16.x builds)
function FindEntryOfficeTicketCache {
	/usr/bin/security find-generic-password -l 'Microsoft Office Ticket Cache' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Ticket Cache' entry from the keychain (16.x builds)
function RemoveEntryOfficeTicketCache {
	/usr/bin/security delete-generic-password -l 'Microsoft Office Ticket Cache' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'com.microsoft.adalcache' entry is present in the keychain (16.x builds)
function FindEntryAdalCache {
	/usr/bin/security find-generic-password -l 'com.microsoft.adalcache' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'com.microsoft.adalcache' entry from the keychain (16.x builds)
function RemoveEntryAdalCache {
	/usr/bin/security delete-generic-password -l 'com.microsoft.adalcache' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Microsoft Office Data' entry is present in the keychain (16.45+ builds)
function FindEntryOfficeData {
	/usr/bin/security find-generic-password -G 'Microsoft Office Data' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Microsoft Office Data' entry from the keychain (16.45+ builds)
function RemoveEntryOfficeData {
	/usr/bin/security delete-generic-password -G 'Microsoft Office Data' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'com.helpshift.data_com.microsoft.Outlook' entry is present in the keychain
function FindEntryHelpShift {
	/usr/bin/security find-generic-password -l 'com.helpshift.data_com.microsoft.Outlook' 2> /dev/null 1> /dev/null
	echo $?
}

# Checks to see if the 'com.helpshift.data_com.microsoft.Outlook' entry is present in the keychain
function FindEntryHelpShift {
	/usr/bin/security find-generic-password -l 'com.helpshift.data_com.microsoft.Outlook' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'com.helpshift.data_com.microsoft.Outlook' entry from the keychain
function RemoveEntryHelpShift {
	/usr/bin/security delete-generic-password -l 'com.helpshift.data_com.microsoft.Outlook' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'MicrosoftOfficeRMSCredential' entry is present in the keychain
function FindEntryRMSCredential {
	/usr/bin/security find-generic-password -l 'MicrosoftOfficeRMSCredential' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'MicrosoftOfficeRMSCredential' entry from the keychain
function RemoveEntryRMSCredential {
	/usr/bin/security delete-generic-password -l 'MicrosoftOfficeRMSCredential' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'com.microsoft.office.rmscache' entry is present in the keychain
function FindEntryRMSCache {
	/usr/bin/security find-generic-password -l 'com.microsoft.office.rmscache' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'com.microsoft.office.rmscache' entry from the keychain
function RemoveEntryRMSCache {
	/usr/bin/security delete-generic-password -l 'com.microsoft.office.rmscache' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'MSProtection.framework.service' entry is present in the keychain
function FindEntryMSProtection {
	/usr/bin/security find-generic-password -l 'MSProtection.framework.service' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'MSProtection.framework.service' entry from the keychain
function RemoveEntryMSProtection {
	/usr/bin/security delete-generic-password -l 'MSProtection.framework.service' 2> /dev/null 1> /dev/null
}

# Checks to see if the 'Exchange' entry is present in the keychain
function FindEntryExchange {
	/usr/bin/security find-generic-password -l 'Exchange' 2> /dev/null 1> /dev/null
	echo $?
}

# Removes the 'Exchange' entry from the keychain
function RemoveEntryExchange {
	/usr/bin/security delete-generic-password -l 'Exchange' 2> /dev/null 1> /dev/null
}

# Evaluate command-line arguments
if [[ $# = 0 ]]; then
	ShowUsage
else
  # If running under Jamf there will be three arguments provided by Jamf, check for user provided --Jamf argument.
  if [[ $# -gt 3 ]]; then
    for KEY in "$@"
    do
      if [[ "$KEY" =~ ^(--jamf|--Jamf)$ ]]; then
        shift 3 # Running under Jamf, ignore first three arguments
      fi
    done
  fi
	for KEY in "$@"
	do
	case $KEY in
    	--Help|-h|--help)
    	ShowUsage
    	shift # past argument
    	;;
    	--Default|-d|--default)
    	REMOVEDEFAULT=true
    	shift # past argument
	;;
    	--IRM|-i|--irm)
    	REMOVEIRM=true
    	shift # past argument
	;;
    	--All|-a|--all)
    	REMOVEALL=true
    	shift # past argument
	;;
    	--Jamf|-j|--jamf)
    	JAMF=true
    	shift # past argument
	;;
	--MAS|-m|--mas)
    	MASCONVERT=true
    	shift # past argument
	;;
	--Force|-f|--force)
    	FORCE=true
    	shift # past argument
	;;
	'') #ignore empty string arguments passed by Jamf
    	shift # past argument
	;;
	*)
    	ShowUsage
    	;;
    esac
	shift # past argument or value
	done
fi

## Main
# Start Removal Routine
APPSRUNNING=$(CheckRunning)
if [ "$APPSRUNNING" == "1" ]; then
	if [ $FORCE ]; then
		ForceQuit2016
	else
		echo "ERROR: Office 2019/2016 apps are open. Either close the apps, or use the '--Force' option."
		echo
		exit 1
	fi
fi

if [ $REMOVEDEFAULT ] || [ $REMOVEALL ] || [ $MASCONVERT ]; then	
	# Find and remove 'msoCredentialSchemeADAL' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeADAL)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryMsoCredentialSchemeADAL
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeADAL)
	done
	
	# Find and remove 'msoCredentialSchemeLiveId' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeLiveId)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryMsoCredentialSchemeLiveId
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeLiveId)
	done

	# Find and remove 'msoCredentialSchemeUnknown' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeUnknown)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryMsoCredentialSchemeUnknown
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryMsoCredentialSchemeUnknown)
	done

	# Find and remove 'com.microsoft.adalcache' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryAdalCache)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryAdalCache
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryAdalCache)
	done
	
	echo "Credential keychain entries removed"
fi

if [ $REMOVEDEFAULT ] || [ $REMOVEALL ]; then	
	# Find and remove 'MSOpenTech.ADAL.1*' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryMSOpenTechADAL1)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryMSOpenTechADAL1
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryMSOpenTechADAL1)
	done
	
	# Find and remove 'Microsoft Office Identities Cache 2' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeIdCache2)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeIdCache2
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeIdCache2)
	done
	
	# Find and remove 'Microsoft Office Identities Cache 3' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeIdCache3)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeIdCache3
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeIdCache3)
	done
	
	# Find and remove 'Microsoft Office Identities Settings 2' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeIdSettings2)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeIdSettings2
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeIdSettings2)
	done
	
	# Find and remove 'Microsoft Office Identities Settings 3' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeIdSettings3)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeIdSettings3
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeIdSettings3)
	done
	
	# Find and remove 'Microsoft Office Ticket Cache' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeTicketCache)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeTicketCache
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeTicketCache)
	done

	# Find and remove 'Microsoft Office Data' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryOfficeData)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryOfficeData
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryOfficeData)
	done

	# Find and remove 'Exchange' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryExchange)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryExchange
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryExchange)
	done

	echo "Default keychain entries removed"
fi

if [ $REMOVEIRM ] || [ $REMOVEALL ]; then	
	# Find and remove 'MicrosoftOfficeRMSCredential' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryRMSCredential)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryRMSCredential
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryRMSCredential)
	done

	# Find and remove 'com.microsoft.office.rmscache' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryRMSCache)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryRMSCache
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryRMSCache)
	done
	
	# Find and remove 'MSProtection.framework.service' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryMSProtection)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryMSProtection
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryMSProtection)
	done
	echo "Rights Management keychain entries removed"
fi

if [ $REMOVEALL ] || [ $MASCONVERT ]; then
	# Find and remove 'com.helpshift.data_com.microsoft.Outlook' entries
	MAXCOUNT=0
	KEYNOTPRESENT=$(FindEntryHelpShift)
	while [ "$KEYNOTPRESENT" == "0" ] || [ $MAXCOUNT -gt 20 ]; do
		RemoveEntryHelpShift
		let MAXCOUNT=MAXCOUNT+1
		KEYNOTPRESENT=$(FindEntryHelpShift)
	done
	echo "HelpShift keychain entries removed"
fi

echo
exit 0