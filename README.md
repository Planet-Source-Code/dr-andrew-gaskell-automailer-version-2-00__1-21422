<div align="center">

## Automailer version 2\.00


</div>

### Description

This code is ideal for sending a large number of formatted e-mails to a distribution list.

Mark Wilson's submission of Automailer v1.0 inspired me to write this. It uses the same principle of sending e-mails to addresses in a database. I re-wrote the project to use CDO rather than MAPI controls. I have extended the code significantly in the following ways:

1) You can mail using RTF (not possible with the MAPI Controls).

2) Full error handling for MAPI errors eg logon failure.

3) Code will stop at EOF in the db, not continue in a loop as in Mark's.

4) No need to write your messages as HTML.

5) No need to have Outlook running for this to work
 
### More Info
 
The zip archive includes an Acces Database, containing 2 fields. Add e-mail addresses to the 1st field, and a flag of 1(send) or 0 (don't send) to the second field.

NB The MS MAPIRTF.DLL file is needed. This is an official DLL from the MS web site. I have included it in the zip archive. Copy this file to your Windows\System directory (for Windows 95) or the Winnt\System32 directory (for

Windows NT.

MAPIRTF includes - "writertf" writes RTF formatted text to a message created by Active Messaging, when passed the profile name, entry ID of the message, entry ID of the message store, and the RTF text. Also contains "readrtf" (not used here)


<span>             |<span>
---                |---
**Submitted On**   |2001-03-01 16:55:40
**By**             |[Dr\. Andrew Gaskell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dr-andrew-gaskell.md)
**Level**          |Intermediate
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD15650312001\.zip](https://github.com/Planet-Source-Code/dr-andrew-gaskell-automailer-version-2-00__1-21422/archive/master.zip)








