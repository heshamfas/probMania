<%@ LANGUAGE="VBSCRIPT" %>
<%
 Option Explicit
 Dim ranWizard, intID, i, background, theme, frameHeight, pagePicture, pageText, pageType, pagewords

 If myinfo.Theme <> "" Then
	Theme = myinfo.Theme
%>
<!--
	$Date: 9/05/97 11:21a $
	$ModTime: $
	$Revision: 8 $
	$Workfile: default.asp $
-->
	<HTML>
	<HEAD>
	<!--#include virtual ="/iissamples/homepage/sub.inc"-->
	<META NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
	<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% call Title %></TITLE>
	<% call styleSheet %>
	<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/mygaminglayout/myga1111.css"><meta name="Microsoft Theme" content="mygaminglayout 1111, default">
<meta name="Microsoft Border" content="b, default">
</HEAD>
	<BODY TopMargin="0" Leftmargin="0"><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">
	<!--#include virtual ="/iissamples/homepage/theme.inc"-->
	<!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<h6>Send mail to <a href="mailto:webmaster@faslabs.com"><font color="#FFFF00">webmaster@FasLabs.com</font></a>  with questions or
comments about this web site.<br>



Last modified: August 31, 2000</h6>



<h5>

Copyright ©2000 FasLabs.  All rights reserved.

</h5>

</td></tr><!--msnavigation--></table></BODY>
	</HTML>
<%
 Else
	response.redirect "/IISSamples/Default/welcome.htm"
 End If
%>