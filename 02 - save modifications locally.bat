@ECHO OFF
SET /P comment=Please enter your comment: 
IF "%comment%"=="" GOTO Error
	git commit -m "%comment%"
GOTO End
:Error
ECHO You did not enter any comment! Bye bye!!
:End
@pause