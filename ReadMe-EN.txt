BuildsHistory.js reads all subkeys like "Source OS*" from HKLM\SYSTEM\Setup. Then for each subkey
it reads most of the values which contains raw data presenting Windows builds properties for each build
ever installed. On base of these data script builds log file in the directory of execution with a 
name BuildsHistory.log. As soon as log file created, script starts notepad.exe to present log and
terminates. Script was not yet tested at 32-bit system but in a case of the error Error.log will
be created which will help to fix an error provided being sent to developer. e-mail is given in the head 
lines of the BuildsHistory.js.

