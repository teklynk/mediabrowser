<!--#include file="../db/dbcon.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--
Filemanager
Created by: Ryan Jones
Date: Nov. 16, 2014
Version: 1.0
www.teklynk.com
-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>MediaBrowser | Welcome</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
<link rel="stylesheet" type="text/css" href="css/bootstrap-override.css" />
<link rel="stylesheet" type="text/css" href="css/styles.css" />
<style type="text/css">
  body {
	padding-top: 60px;
	padding-bottom: 40px;
  }
  .sidebar-nav {
	padding: 9px 0;
  }

  @media (max-width: 980px) {
	/* Enable use of floated navbar text */
	.navbar-text.pull-right {
	  float: none;
	  padding-left: 5px;
	  padding-right: 5px;
	}
  }
</style>
<link rel="stylesheet" type="text/css" href="css/bootstrap-responsive.min.css" />
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/bootstrap.min.js"></script>

<!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
<!--[if lt IE 9]>
  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
  <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
<![endif]-->

</head>
<body>
<%

set dbconn = server.createobject("ADODB.Connection")
dbconn.open DBName

'''varialbles used through out the site
pagename = Lcase(Request.ServerVariables("SCRIPT_NAME"))

if inStr(pagename,"default.asp")=0 then%>
<div class="navbar navbar-inverse navbar-fixed-top">
  <div class="navbar-inner">
	<div class="container-fluid">
<!--Navigation Bar-->
	<!-- Responsive Navbar Part 1: Button for triggering responsive navbar (not covered in tutorial). Include responsive CSS to utilize. -->
	<button type="button" class="btn btn-navbar" data-toggle="collapse" data-target=".nav-collapse">
	  <span class="icon-bar"></span>
	  <span class="icon-bar"></span>
	  <span class="icon-bar"></span>
	</button>
	<a class="brand" href="filemgr.asp">MediaBrowser</a>
	<!-- Responsive Navbar Part 2: Place all navbar contents you want collapsed withing .navbar-collapse.collapse. -->
	<div class="nav-collapse collapse">
		<p class="navbar-text pull-right">
		  Logged in as <%=Session("username")%>
		  <i class="divider-vertical"></i>
		  <a href="filemgr.asp?logoff=1">Sign Out</a>
		</p>
	  <ul class="nav">

		<li class="dropdown">
		  <a href="?myplaylist=true">My Playlist</a>
		</li>

	  </ul>
	  </div><!--/.nav-collapse -->
	</div>
  </div>
</div>
<%end if%>
<!--Main Body-->
<div class="container">