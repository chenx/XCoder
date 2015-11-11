XCoder
======

XCoder is a website framework generator written in VB.NET. It generates a website framework in ASP.

XCoder was written back in 2004. It summarizes several patterns of IMS (Information Management System), and generates template based these patterns. This can immensely speed up the development speed. 

This can potentially be extended to more patterns, and generate code in other languages.

For more details, see Help.chm (extracted from <a href="https://github.com/chenx/XCoder/releases/download/XCoder_v1.0.0_2005/Help.zip">Help.zip</a>).

<h2>1. Overview</h2>

In short, XCoder is a RAD (Rapid Application Development) tool. It is used to help developing database driven website.

A little more specific, XCoder is a code generator. It automatically generates code for a database driven website or part of it. The website developer then can customize the code. The automatic generation part can save over 90% of the development time and allow the developer to concentrate on high level function design instead of low level programming details. Figure 1 shows what XCoder does. 

<p align="center">
<img src="http://cssauh.com/xc/demo/XCoder/README/image/fig1.jpg">
<br/>Figure 1. Role of XCoder
</p>

<h2>2. Detail</h2>

<h3>2.1 A little background.</h3>

Current popular programming languages for creating database driven website include ASP, ASP.NET, PHP, JSP and CodeFusion etc. ASP was released by Microsoft in 1996. XCoder generates code in ASP.

<h3>2.2 XCoder as a web application.</h3>

This is the early version of XCoder written in the fall of 2003 in my Research Assistant work for the MarBEC center of University of Hawaii. It is part of the <a href="http://www.hawaii.edu/epscor/">EPSCoR</a> website and helped with its development. It generates an information management module for a website, which allows information input, edit, view and delete. Figure 2 is a screenshot of the application in action.

<p align="center">
<img src="http://cssauh.com/xc/demo/XCoder/README/image/fig2.jpg">
<br/>Figure 2. Screenshot of XCoder web version
</p>

<h3>2.3 XCoder as a windows application.</h3>

This is the new version of XCoder which I started this year. It is totally rewritten in VB.NET and has many new features. It can generate a website framework that allows user to sign up and log in. Within this framework these functional modules can be generated: information management (the function of the web version above), bulletin board, file upload, search, calendar, data export to Word, Excel and Access. Figure 3 is a screenshot of the software. 

<p align="center">
<img src="http://cssauh.com/xc/demo/XCoder/README/image/fig3.jpg">
<br/>Figure 3. Screenshot of XCoder desktop version
</p>

<h2>3. Prospect</h2>

There is always a big demand for database driven website. Although new technologies like ASP.NET and J2EE provide stronger support for complicated web applications, the cost/price ratio of XCoder and its ease of use are promising in getting its market share. XCoder has the potential of being extended to generate code in other programming languages. New module templates can be added. Customization and specialization are possible. 

<h2>4. Installer</h2>

<a href="http://cssauh.com/xc/demo/XCoder/">XCoder Application Installer (XCoder Beta Setup.msi)</a>.

Mahalo!


Source code and Installer
=========================

This repository contains the following files: 

src/ - contains the source code.     
src/components/ - contains the VB code to generate GetSQLServers.dll. This dll returns a list of available DSN connections on local machine.    
src/XCoder_2004_07-12_src - the source code as of 7/11/2004    
setup/ - the XCoder installer source and output. This is what you use to install XCoder on your windows machine.    

System requirement
==================

XCoder was developed in Visual Studio.NET 2003. 

The system requirements to install and use XCoder are:    
- windows 2000 or above    
- .NET framework 1.1 or above    
- and MSSQL 7.0 or above   

Author
======
X. Chen, Copyright @ 2004 - 2005


