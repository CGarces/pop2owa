---
layout: default
title: Preguntas Frecuentes
name: faq
---

## Contenido

*   [General](/es/index.php?page=FAQ.html#General)
    *   [El correo aparece sin acentos ni eñes](/es/index.php?page=FAQ.html#UTF8)
    *   [Los correos que envío se reciben con caracteres extraños](/es/index.php?page=FAQ.html#quoted_printable)
*   [Thunderbird](/es/index.php?page=FAQ.html#General)
    *   [Las imágenes no aparecen en algunos mensajes](/es/index.php?page=FAQ.html#Images_in_messages_do_not_appear)

## General

<a name="UTF8"></a>

### El correo aparece sin acentos ni eñes

Outlook puede enviar (especialmente en mensajes HTML un mensaje con una cabecera incorrecta (UFT-8), para leer correctamente el mensaje es necesario cambiar la codificación del mensaje (en castellano ISO-8859-1)<a name="quoted_printable"></a>

### Los correos que envío se reciben con caracteres extraños

(O)utlook (W)eb (A)ccess debe recibir los mensajes en formato de 7bits. Se debe forzar el envío de los mensajes codificados como "entre comillas" (quoted_printable) Si usas outlook Express, la configuración correcta se indica en el [tutorial](/es/index.php?page=Tutorial/outlook.html#quoted_printable)  
En thunderbird 1.5 la opción esta disponible en Herramientas->Opciones->Redacción Mas información sobre el problema en el [foro de desarrolladores](http://sourceforge.net/forum/forum.php?thread_id=1587883&forum_id=508559)  

## Thunderbird

<a name="Images_in_messages_do_not_appear"></a>

### Las imágenes no aparecen en algunos mensajes

Los mensajes compuestos con word/outlook pueden dar problemas, en el caso del outlook la causa más común es en envío de la cabecera incorrecta en las imágenes (Content-Type: _application/octet-stream_ en vez de _image/jpeg_)  
Más información en el articulo correspondiente de [MozillaZine Knowledge Base](http://kb.mozillazine.org/Images_in_messages_do_not_appear)
