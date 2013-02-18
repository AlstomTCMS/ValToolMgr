cd ..
@ECHO OFF
SET /P branch=Please enter branch (optional): 
IF "%branch%"=="" GOTO without_branch
	git pull -v --progress origin %branch%
GOTO End
:without_branch
	git pull -v --progress origin
:End
@pause