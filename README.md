# SalvaRed

General system description
=======================

This program does a simple cycle checking for changes in a folder in the file system and saving them to a backup location in a central server.
At start, it connects to the central server and reads the structure of the backup folder storing it in a local memory structure for that purpose. From there on, it checks for changes periodically comparing the local file system folder to the memory copy of the remote save. When a change is detected (file comparison on modified date) the remote file is updated (as well as it's image in the local memory structure).

It's development took part around 2006 ~ 2007 and was used in production without any glitch in a small network of around 16 to 20 computers for a year or so.

That functionallity was needed again in 2015 on a similar network but with a greater count of computers (around 50 or 60). For this ocasion a reimplementation was made as a console program in C++ using visual studio. That reimplementation was made as a different project, with a similar aproach though.



