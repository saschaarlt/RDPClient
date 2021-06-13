# RDPClient (console)
This is the RDP Client Console which is used to automaticly connects to a RDP Session.


We're looking for a rdp login class written in vb.net. The class would be possible to login via RDP with a specified user and a given file. Returning a possible login or failure (with rdp message)
Main Features:
Login / RDP Login execution with a given (unsigned) connection.rdp file
Check if Login was possible
Returning error, the rdp exception message
NO GUI! All running unattended!
RDP Logoff, correct User Logoff from RDP Session
TimeSpan for Login
Test Driven Development and Delivery
We prefer Test Driven development for this order so we start writing tests for the functionality we would order. At the start of the order the following things are described in the attached UnitTest.

The class has to do the following:
+ Start a RDP Connection based on the attached cpub-TSFARMDEV-CmsRdsh.rdp
+ Login with a specified user (not in RDP) and a specified password (both were given when you get the order)
+ Wait for the Login / RDP Connection is established
+ optional: User Logoff from RDP Session via  a alternate shell:s:"C:\topsecret\WaitAndLogoff.cmd" in RDP Config
+ return the connection was successful in a variable
+ return that the connections fails with the correct RDP Error message
Notice
The class would also work with different cpub-TSFARMDEV-CmsRdsh.rdp files.
It is planned to extend the class later with additional orders. This is a first release candidate but must be fully functional and perform all tests.
Assistance
The Post https://www.codeproject.com/Articles/33979/Multi-RDP-Client-NET-A-Simple-and-Friendly-User-In describe all of our needs and is available as a github Open Source project https://github.com/jaysonragasa/MultiRDPClient.NET. We believe that a large part of the task can be done quickly and easily. Please take a look at it.
Coding defaults:
You can code you like but the delivery must be in VB.NET!!
object oriented, maybe classes and structures where it was meaningful
Maybe result of connection in a structure / result class
Code comment on possible difficult places
Crash before success. The catch of error is especially important.
Estimates:
Delivery:
Code only in VB.net! Delivered by a Github repository (checkin) that's earmarked by us (private). You also need a free github account and knowledge how to check code in (push) and get out (pull) our UnitTest cloned from Repo.
Finished:
1. Revision with the coordinator (me) with a fully UnitTest of all functions.
2. by a Developer of us, which investigate the code and my reconsider some code or extend the UnitTests
3. When all UnitTests and actions are successful happened
Please offer 1. your price, 2. delivery date (days) and 3. skill (if not in your profile), independent from our budget. Please inform us if you’re doing something like this before and / or have special knowledge about rdp.
