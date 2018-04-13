echo off

REM Macro batch processing - FactSage 6.3

REM Run_Equilib.bat & Run_PhaseDiagram.bat are not supported on Client installations 
REM - you must use a FactSage Standalone Installation (PC has a dongle). 

REM A record of the mac processing is written to the file macro.log 

REM The following command loads Equilib and runs the macro file Macros\Equi....

EquiSage.exe /EQUILIB /MACRO  D:\ExpertSystem\ISA.mac
