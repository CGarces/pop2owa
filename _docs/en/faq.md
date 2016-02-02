---
layout: default
title: FAQ
name: faq
---

## General

### I have encoding problems in incoming mails with non US charset

Outlook send some messages (some HTML emails) with incorrect charset headers (UFT-8), please manually change the charset from your email client.<a name="quoted_printable"></a>

### I have encoding problems in outgoing mails with non US charset

(O)utlook (W)eb (A)ccess must send the mails in quoted_printable format. Please force this configuration in your e-mail client. If you use Outlook Express, please read the [tutorial]({{"/tutorial_pop2owa1/" | prepend: site.baseurl }})  
In Thunderbird you can force it under "Composition" options, "For messages that contain 8-bit characters, use 'quoted printable' MIME encoding..."  
More info about this issue at [developer forum](http://sourceforge.net/forum/forum.php?thread_id=1587883&forum_id=508559)  

## Thunderbird

### Images on messages not appear under some emails clients

Outlook can send images in HTML mails with incorrect headers. In version 1.1.7 the headers are parsed to try to fix this error.  
More info about this issue at [MozillaZine Knowledge Base](http://kb.mozillazine.org/Images_in_messages_do_not_appear)
