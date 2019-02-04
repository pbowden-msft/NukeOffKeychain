# NukeOffKeychain
<b>Microsoft Office 365/2019/2016 Keychain Removal Tool</b>

Purpose: Removes Office 365/2019/2016 for Mac keychain entries</br>
Usage: `NukeOffKeychain [--Default] [--IRM] [--All] [--Force] [--Jamf]`</br>
&nbsp;&nbsp;&nbsp;`[--Default]` removes the logon, cache and ADAL keychain entries</br>
&nbsp;&nbsp;&nbsp;`[--IRM]` removes the rights management keychain entries</br>
&nbsp;&nbsp;&nbsp;`[--All]` removes logon, cache, ADAL, rights management and HelpShift keychain entries</br>
&nbsp;&nbsp;&nbsp;`[--Jamf]` ignores the first 3 arguments for running via Jamf</br>
&nbsp;&nbsp;&nbsp;`[--MAS]` removes the entries required for a clean CDN to Mac App Store conversion</br>
