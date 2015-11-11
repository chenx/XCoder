XCoder
======

XCoder is a website framework generator written in VB.NET. It generates a website framework in ASP.

XCoder was written back in 2004. It summarizes several patterns of IMS (Information Management System), and generates template based these patterns. This can immensely speed up the development speed. In my case, it cuts down 80% of development time and cost.

This can potentially be extended to more patterns, and generate code in other languages.

For more details, see Help.chm (extracted from <a href="https://github.com/chenx/XCoder/releases/download/XCoder_v1.0.0_2005/Help.zip">Help.zip</a>).

<h3>1. Overview</h3>

In short, XCoder is a RAD (Rapid Application Development) tool. It is used to help developing database driven website.

A little more specific, XCoder is a code generator. It automatically generates code for a database driven website or part of it. The website developer then can customize the code. The automatic generation part can save over 90% of the development time and allow the developer to concentrate on high level function design instead of low level programming details. Figure 1 shows what XCoder does. 


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


