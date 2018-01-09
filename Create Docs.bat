REM del /Q help\ANSI\Doxygen\public\html\search\*
REM del /Q help\ANSI\Doxygen\public\html\*
REM rd help\ANSI\Doxygen\public\html\search
REM del /Q help\Unicode\Doxygen\private\html\search\*
REM del /Q help\Unicode\Doxygen\private\html\*
REM del /Q help\Unicode\Doxygen\public\html\search\*
REM del /Q help\Unicode\Doxygen\public\html\*
REM 
REM doxygen "help\ANSI\Doxygen\public docs CHM.cfg"
REM doxygen "help\Unicode\Doxygen\private docs CHM.cfg"
REM doxygen "help\Unicode\Doxygen\public docs CHM.cfg"

del /Q help\ANSI\Doxygen\public\html\search\*
del /Q help\ANSI\Doxygen\public\html\*
rd help\ANSI\Doxygen\public\html\search
del /Q help\Unicode\Doxygen\private\html\search\*
del /Q help\Unicode\Doxygen\private\html\*
del /Q help\Unicode\Doxygen\public\html\search\*
del /Q help\Unicode\Doxygen\public\html\*

doxygen "help\ANSI\Doxygen\public docs.cfg"
doxygen "help\Unicode\Doxygen\private docs.cfg"
doxygen "help\Unicode\Doxygen\public docs.cfg"
