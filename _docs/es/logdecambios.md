---
layout: default
title: Log de Cambios
name: Changelog
---

**Log de cambios**

Cambios en la versión v.1.3.6  

*   Solucionado error en mensajes con la secuencia CRLF+".".+CRLF [Bug [2220343](http://sourceforge.net/tracker2/?func=detail&aid=2220343&group_id=152204&atid=783596)]

Cambios en la versión v.1.3.5  

*   Corregido error enviando cabeceras de cookies. (Gracias a li0n40) [Bug [2080904](http://sourceforge.net/tracker/index.php?func=detail&aid=2080904&group_id=152204&atid=783596)]
*   Corregido error en el tratamiento de adjuntos.

Cambios en la versión v.1.3.4  

*   Corregido error en el tratamiento de adjuntos. [Bug [2031676](http://sourceforge.net/tracker/index.php?func=detail&aid=2031676&group_id=152204&atid=783596)]
*   Corregido error con el comando AUTH PLAIN. [Bug [2000676](http://sourceforge.net/tracker/index.php?func=detail&aid=2000676&group_id=152204&atid=783596)]
*   Corregido error menor con commandos TOP y LIST.
*   Modificadas respuestas del servidor SMTP para adecuarse al estandar.

Cambios en la versión v.1.3.3  

*   Corregido error en login mediante formulario. [Bug [1942104](http://sourceforge.net/tracker/index.php?func=detail&aid=1942104&group_id=152204&atid=783596)]
*   Cambios menores en código fuente.

Cambios en la versión v.1.3.2  

*   Añadido codigo para mantener el objeto de conexion anterior (el de la rama estable), algunos usuarios no pueden conectarse con el nuevo
*   Corregido bug con la etiqueta "Hidden_Field" en el fichero de configuracion
*   Cambios menores en código fuente.

Cambios en la versión v.1.3.1  

*   Mejorado el soporte para proxys
*   Solucionado bug en instalador. [Bug [1916933](http://sourceforge.net/tracker/index.php?func=detail&aid=1916933&group_id=152204&atid=783596)]
*   Cambios menores en código fuente.

Cambios en la versión v.1.3.0  

*   Añadido soporte para certificados SSL
*   Valores de configuración movidos a fichero XML.

Cambios en la versión v.1.2.2  

*   Corregido error al enviar parametros en la autentificación por formulario. [Bug [1867273](http://sourceforge.net/tracker/index.php?func=detail&aid=1867273&group_id=152204&atid=783596)] (gracias a jcdailey)
*   Añadido código para evitar el envio/recepción de mensajes simultaneo.

Cambios en la versión v.1.2.1  

*   Corregido error en el formateo de una URL, introducido en la version 1.2.
*   Mejorada la gestión de cabeceras de los mails.

Cambios en la versión v.1.2  

*   Corregidos errores ortograficos en los mensajes.
*   Cambiado código fuente para usar el nuevo fichero de configuración.
*   Añadido código para resetear la conexion cuando esta es cancelada por el cliente de correo.
*   Cambios menores el las llamadas a las APIs de windows.
*   Corregido error menor en Windows XP.
*   Corregido error recuperando configuracion SMTP.
*   Mejorado el log y la gestion de errores.

Cambios en la versión v.1.2 RC1  

*   Lenguage por defecto del instalador cambiado a ingles [Bug [1770838](http://sourceforge.net/tracker/index.php?func=detail&aid=1770838&group_id=152204&atid=783596)].
*   Corregido error en el folor de fondo de las etiquetas.
*   Corregidos errores ortográficos.
*   Añadido código para corregir codificaciones incorrectas de los ficheros GIF.
*   Corregido error menor con servidores Exchange 2003.
*   Eliminado código no usado.
*   Mejorada la documentació del código fuente.

Cambios en la versión v.1.1.7  

*   Corregido error de cache en las peticiones GET [[1756959](http://sourceforge.net/tracker/index.php?func=detail&aid=1756959&group_id=152204&atid=783596)].
*   Corregido error cuando se envían y reciben correos simultáneamente.
*   Añadido código para soportar cuentas con el formato DOMAIN\User^Mailbox [[1739893](http://sourceforge.net/tracker/index.php?func=detail&aid=1739893&group_id=152204&atid=783596)].
*   Añadido código para corregir codificaciones incorrectas de los ficheros JPG.
*   Añadido código adicional para evitar un posible error al enviar mensajes en servidores con la autentificación por formulario activa.
*   Añadido código para manejar cadenas en Base64 usando MSXML.
*   Cambiada la petición de GetMsgList para reducir el tamaño de la respuesta.
*   Aumentado el tamaño del buffer.
*   Eliminado código no usado.
*   Cambios menores en código fuentes.
*   Comentarios modificados al formato de VBDOX.

Cambios en la versión v.1.1.6  

*   Añadido control mscomctl.ocx al instalador. Error [[174252](http://sourceforge.net/tracker/index.php?func=detail&aid=174252&group_id=152204&atid=783596)]
*   Implementado parámetro -quit para terminar el proceso. Error [[1741773](http://sourceforge.net/tracker/index.php?func=detail&aid=1741773&group_id=152204&atid=783596)]
*   Añadido código adicional para permitir solo una instancia en ejecución.

Cambios en la versión v.1.1.5  

*   Cambiada la autentificación por formulario, para soportar servidores Exchange 2007\. Bug [[1619844](http://sourceforge.net/tracker/index.php?func=detail&aid=1741769&group_id=152204&atid=783596)] (gracias a f.hartmann)
*   Mejorado el log para dar mejor soporte en caso de problemas o errores.
*   Modificado el código de autentificación STMP.
*   Corregido error cuando se introduce una URL terminada en "/"

Cambios en la versión v.1.1.4  

*   Corregido error [[1619844](http://sourceforge.net/tracker/index.php?func=detail&aid=16198448&group_id=152204&atid=783596)] en la conversión de caracteres UNICODE/ANSI (gracias a Aleksey Pershin)
*   Corregido error [[1621688](http://sourceforge.net/tracker/index.php?func=detail&aid=1621688&group_id=152204&atid=783596)] con cabeceras mayores de 32KB
*   Corregido error en servidores con buzones del tipo user@company.com y la autentificación por formulario activa
*   Cambios menores en código fuente

Cambios en la versión v.1.1.3  

*   Código de manejo de sockets movido a la clase CSocketMaster, parece mas estable y quizas elimine los bloqueos del GUI
*   Corregido error en el formateo de direcciones de e-mail

Cambios en la versión v.1.1.2  

*   Mejorada la gestión de errores y su log

Cambios en la versión v.1.1.1  

*   código para obtener lo mensajes totalmente reescrito, puede ser inestable.
*   Añadidas opciones de linea de comandos

*   -v 0 a -v 3 nivel de detalle del log
*   -NT para ejecutarse como servicio NT

Cambios en la versión v.1.1  

*   Corregido parcialmente el error con correos adjuntos [[bug 1581048](http://sourceforge.net/tracker/index.php?func=detail&aid=1581048&group_id=152204&atid=783596)]
*   Pop2owa puede ejecutarse como servicio NT (lee [esto](http://support.microsoft.com/kb/q137890/)) usando la cuenta del usuario logado

Cambios en la versión v.1.0  

*   Solucionado error [[1417370](http://sourceforge.net/tracker/index.php?func=detail&aid=1417370&group_id=152204&atid=783596)] con adjuntos grandes

Cambios en la versión v.1.0 RC5  

*   Fixed bug [[1569922](http://sourceforge.net/tracker/index.php?func=detail&aid=1569922&group_id=152204&atid=783596)] Port number settings are not saved
*   Minor changes in GUI and error handler

Cambios en la versión v.1.0 RC4  

*   Changed HTTPXML commands to asynchronous to fix GUI locks [[1544517](http://sourceforge.net/tracker/index.php?func=detail&aid=1544517&group_id=152204&atid=783596)].
*   Fixed minor bug with multi-part messages.

Cambios en la versión v.1.0 RC3  

*   Fixed bug with Non Delivery Reports [[1494455](http://sourceforge.net/tracker/index.php?func=detail&aid=1494455&group_id=152204&atid=783596)].
*   Added code to retry if send fails.

Cambios en la versión v.1.0 RC2  

*   Fixed bug with attachments if the msg not is multipart.
*   Striped the final dot in send messages.
*   Increased performance sending mails.

Cambios en la versión v.1.0 RC  

*   Added systray feature
*   Optimized memory usage
*   Fixed GUI freeze problems

Cambios en la versión v.0.11  

*   Better error handling
*   Less memory usage

Cambios en la versión v.0.10  

*   SMTP authentication re-writed.

Cambios en la versión v.0.9  

*   Fixed bug with Encoded Content-Transfer-Encoding [1458243].
*   Rewrited attachment headers parsing.

Cambios en la versión v.0.8  

*   Added code to support Form Autentication. Now you can use POP2OWA in a server with Form-Based-Authentication turned on.
*   Added message size un STAT, LIST and RETR commands conform to the standard for the format of Internet text messages [RFC822].

Cambios en la versión v.0.7.1  

*   Dump all errors in a log (pop2owa.err) to help to fix the bugs.

Cambios en la versión v.0.7  

*   Fixed bug with HTML messages without attachments.
*   The source code has been documented with [VBDOXAddin](http://sourceforge.net/projects/vbdoxaddin).

Cambios en la versión v.0.6  

*   Internal code has re-writed to put POP3 and WebDav code in separate classes.
*   Some minor bugs fixed

Cambios en la versión v.0.5  

*   Fixed error in Content-Type header [1364201]
*   Fixed error in string comparisons [1363892]
*   Installer not overwrite the configuration [1362635]
*   Minor bugs fixed

Cambios en la versión v.0.4  

*   Fixed bug in AUTH command [1356534]
*   Fixed bug in mail priorities
*   Fixed bug in attachment without content type
