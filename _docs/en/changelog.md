---
layout: default
title: Change Log
name: Changelog
---

**Change Log**

Changes from version v.1.3.6  

*   Fixed bug with sequence CRLF+".".+CRLF [Bug [2220343](http://sourceforge.net/tracker2/?func=detail&aid=2220343&group_id=152204&atid=783596)]


Changes from version v.1.3.5  

*   Fixed bug send cookies headers. (Thanks to li0n40) [Bug [2080904](http://sourceforge.net/tracker/index.php?func=detail&aid=2080904group_id=152204&atid=783596)]
*   Fix error parsing attachments.

Changes from version v.1.3.4  

*   Fixed bug parsing attachments. [Bug [2031676](http://sourceforge.net/tracker/index.php?func=detail&aid=2031676&group_id=152204&atid=783596)]
*   Fixed bug with AUTH PLAIN command. [Bug [2000676](http://sourceforge.net/tracker/index.php?func=detail&aid=2000676&group_id=152204&atid=783596)]
*   Fixed minor bug with TOP and LIST commands.
*   SMTP responses changes to RFC standard.

Changes from version v.1.3.3  

*   Fixed bug with FBA login. [Bug [1942104](http://sourceforge.net/tracker/index.php?func=detail&aid=1942104&group_id=152204&atid=783596)]
*   Minor changes in source code.

Changes from version v.1.3.2  

*   Added code to support old connection object (from stable branch), some user have problems with the new object
*   Fixed bug with "Hidden_Field" tags in config file
*   Minor changes in source code.

Changes from version v.1.3.1  

*   Better proxy support
*   Fixed bug in Uninstaller. [Bug [1916933](http://sourceforge.net/tracker/index.php?func=detail&aid=1916933&group_id=152204&atid=783596)]
*   Minor changes in source code.

Changes from version v.1.3.0  

*   Initial SSL support, now pop2owa can works with invalid SSL certificates.
*   Configurations values moved to XML file, now pop2owa not use registry (only the installer use it).

Changes from version v.1.2.2  

*   Fixed error sending parameters in Form Based Authentication. [Bug [1867273](http://sourceforge.net/tracker/index.php?func=detail&aid=1867273&group_id=152204&atid=783596)] (thanks to jcdailey)
*   Added code to avoid send/receive mail at the same time.

Changes from version v.1.2.1  

*   Fixed bug formatting URL introduced in 1.2.
*   Improved header parsing.

Changes from version v.1.2  

*   Fixed typo errors in messages.
*   Changed source code to support new config file.
*   New code to reset the connection when email client close the socket.
*   Minor changes in API calls.
*   Fixed bug in some Windows XP.
*   Fixed error retrieving SMTP configuration.
*   Improved log to detect bugs, better error handling.

Changes from version v.1.2 RC1  

*   Installer default language changed to English [Bug [1770838](http://sourceforge.net/tracker/index.php?func=detail&aid=1770838&group_id=152204&atid=783596)].
*   Fixed minor bug with label backgrounds.
*   Fixed minor typo bug.
*   Fixed incorrect headers in GIF files.
*   Fixed minor bug with 2003 servers.
*   Removed unused code.
*   Improved source code documentation.

Changes from version v.1.1.7  

*   Fixed cache issue in GET request [[1756959](http://sourceforge.net/tracker/index.php?func=detail&aid=1756959&group_id=152204&atid=783596)].
*   Fixed bug if send and receive mails at same time.
*   Added code to support accounts with format DOMAIN\User^Mailbox [[1739893](http://sourceforge.net/tracker/index.php?func=detail&aid=1739893&group_id=152204&atid=783596)].
*   Added code to fix incorrect headers on JPG files.
*   Added additional code in sent function to avoid possible errors on servers with form authentication enabled.
*   Added code to handle Base64 strings using MSXML.
*   Changed GetMsgList request to reduce response size.
*   Increased buffer size.
*   Removed unused code.
*   Minor changes in source code.
*   Comments formatted in VBDOX format

Changes from version v.1.1.6  

*   Added mscomctl.ocx to the installer. Bug [[174252](http://sourceforge.net/tracker/index.php?func=detail&aid=174252&group_id=152204&atid=783596)]
*   Implemented -quit parameter to kill pop2owa process. Bug [[1741773](http://sourceforge.net/tracker/index.php?func=detail&aid=1741773&group_id=152204&atid=783596)]
*   Added additional code to avoid more than one instance of pop2owa.

Changes from version v.1.1.5  

*   Changed Form-Based Authentication function to support Exchange 2007\. Bug [[1619844](http://sourceforge.net/tracker/index.php?func=detail&aid=1741769&group_id=152204&atid=783596)] (thanks to f.hartmann)
*   Improved Log for better support on bugs and problems.
*   Changed STMP authentication code.
*   Fixed bug when URL finish with "/"

Changes from version v.1.1.4  

*   Fixed bug [[1619844](http://sourceforge.net/tracker/index.php?func=detail&aid=1619844&group_id=152204&atid=783596)]with UNICODE/ANSI conversion (thanks to Aleksey Pershin)
*   Fixed bug [[1621688](http://sourceforge.net/tracker/index.php?func=detail&aid=1621688&group_id=152204&atid=783596)] with headers bigger than 32KB
*   Fixed bug with servers with inbox like user@company.com and form based authentication enabled
*   Minor changes in source code

Changes from version v.1.1.3  

*   Socket code changed to CSocketMaster, this make the program more stable and MAYBE fix the GUI problems [[1544517](http://sourceforge.net/tracker/index.php?func=detail&aid=1544517&group_id=152204&atid=783596)]
*   Fix incorrect email address encoding

Changes from version v.1.1.2  

*   Improved application log

Changes from version v.1.1.1  

*   Code to get messages it's re-writed, this code may be unstable, please give me feedback.
*   Command line options added

*   -v 0 to -v 3 verbose level
*   -NT to run as NT service

Changes from version v.1.1  

*   Partial fix of embedded mails [[bug 1581048](http://sourceforge.net/tracker/index.php?func=detail&aid=1581048&group_id=152204&atid=783596)]
*   Pop2owa can run as NT service (read [this](http://support.microsoft.com/kb/q137890/)) using logged user account

Changes from version v.1.0  

*   Fixed bug [[1417370](http://sourceforge.net/tracker/index.php?func=detail&aid=1417370&group_id=152204&atid=783596)] with big attachments

Changes from version v.1.0 RC5  

*   Fixed bug [[1569922](http://sourceforge.net/tracker/index.php?func=detail&aid=1569922&group_id=152204&atid=783596)] Port number settings are not saved
*   Minor changes in GUI and error handler

Changes from version v.1.0 RC4  

*   Changed HTTPXML commands to asynchronous to fix GUI locks [[1544517](http://sourceforge.net/tracker/index.php?func=detail&aid=1544517&group_id=152204&atid=783596)].
*   Fixed minor bug with multi-part messages.

Changes from version v.1.0 RC3  

*   Fixed bug with Non Delivery Reports [1494455]
*   Added code to retry if send fails.

Changes from version v.1.0 RC2  

*   Fixed bug with attachments if the msg not is multipart.
*   Striped the final dot in send messages.
*   Increased performance sending mails.

Changes from version v.1.0 RC  

*   Added systray feature
*   Optimized memory usage
*   Fixed GUI freeze problems

Changes from version v.0.11  

*   Better error handling
*   Less memory usage

Changes from version v.0.10  

*   SMTP authentication re-writed.

Changes from version v.0.9  

*   Fixed bug with Encoded Content-Transfer-Encoding [1458243].
*   Rewrited attachment headers parsing.

Changes from version v.0.8  

*   Added code to support Form Authentication. Now you can use POP2OWA in a server with Form-Based-Authentication turned on.
*   Added message size on STAT, LIST and RETR commands conform to the standard for the format of Internet text messages [RFC822].

Changes from version v.0.7.1  

*   Dump all errors in a log (pop2owa.err) to help to fix the bugs.

Changes from version v.0.7  

*   Fixed bug with HTML messages without attachments.
*   The source code has been documented with [VBDOXAddin](http://sourceforge.net/projects/vbdoxaddin).

Changes from version v.0.6  

*   Internal code has re-writed to put POP3 and WebDav code in separate classes.
*   Some minor bugs fixed

Changes from version v.0.5  

*   Fixed error in Content-Type header [1364201]
*   Fixed error in string comparisons [1363892]
*   Installer not overwrite the configuration [1362635]
*   Minor bugs fixed

Changes from version v.0.4  

*   Fixed bug in AUTH command [1356534]
*   Fixed bug in mail priorities
*   Fixed bug in attachment without content type
