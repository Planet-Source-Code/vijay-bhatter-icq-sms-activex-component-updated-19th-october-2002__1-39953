For all those developers who want to integrate the Mobile SMS feature in their applications I recently developed an ActiveX dll component in Visual Basic which can send your SMS messages through the ICQ Network from ASP / Visual Basic applclications. 

======================
Installation Procedure
======================
1. Compile the code to ICQSMS.dll.

2. Copy the above component in your  C:\Winnt\System32 folder.

3. Register the dll. 

eg. regsvr32 C:\winnt\system32\icqsms.dll

4. Change the ICQ No. and password with yours else the code will not work. If you do not have a ICQ account & password you need to register 1 at http://www.icq.com


======================
Minimum Requirements:
======================
To run this application you'll need Microsoft Internet Transfer Control and Visual Basic Runtime dll installed onto your system. 


======================
Sample ASP Code
======================
<%

set ICQSMS = CreateObject("ICQSMS.SMS")
 
ICQSMS.DeveloperName = "Vijay Bhatter"
ICQSMS.DeveloperEmail = "vijay@cyberax.net"
ICQSMS.ResponseTime = 15
 
'Syntax: 'ICQSMS.SendSMS ICQUIN,Password,CountryCOde,Carrier,MobileNumber,Message
'eg:

Change the ICQ No. and password with yours else the code will not work. If you do not have a ICQ account & password you need to register 1 at http://www.icq.com

ICQSMS.SendSMS "30110446","xdwswe","91", "983", "0144101","I love SMS" 
 
set ICQSMS = nothing
 
%>


======================
Request
======================
I would be glad if you all please do give your valuable feedback and yes don't forget to vote for it. Till then Happy SMS'ing.

