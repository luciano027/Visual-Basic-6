/***
*	RBckLib.CH
*
*	Header file for the RBckLib v1.0d CA-Clipper Library
*
*	Copyright (c) 1993-95  Rolf van Gelder, Eindhoven
*	All rights reserved
*
*	Revision date : 18/01/95
*/

*==========================================================================
*	Return codes for RBckLib functions
*==========================================================================

#define	BU_OKAY			0	&& No errors detected
#define	BU_INVALID_PARM		1	&& Invalid parameter passed
#define	BU_OPEN_INPUT		2	&& Error opening  input  file
#define	BU_WRONG_VERSION	3	&& Wrong version of RBckLib
#define	BU_CREATE_OUTPUT	4	&& Error CREATING output file
#define	BU_READ_INPUT		5	&& Error READING  input  file
#define	BU_WRITE_OUTPUT		6	&& Error WRITING  output file
#define	BU_NO_FILES_FOUND	7	&& No files found to backup
#define	BU_USER_ABORT		8	&& Function aborted by user
#define BU_INVALID_DRIVE  	9	&& Invalid backup drive
#define BU_INVALID_DIR  	10	&& Invalid restore directory

*==========================================================================
*	Array with pre-defined error messages
*==========================================================================
#define	BU_ERRMSG { ;
	"Invalid parameter(s) passed", ;
	"Error OPENING input file", ;
	"Wrong version of RBckLib", ;
	"Error CREATING output file", ;
	"Error READING input file", ;
	"Error WRITING output file", ;
	"No files found to backup", ;
	"Function aborted by user", ;
	"Invalid backup drive/disk", ;
        "Diretorio Invalido para Backp" }

*==========================================================================
*	R_BckList() Subarray Structure
*==========================================================================
#define	BU_FNAME		1	&& Original file name
					&&    (d:\path\filename.ext)
#define	BU_FSIZE		2	&& Original file size
#define	BU_FDATE		3	&& Original file date (dd-mm-yyyy)
#define	BU_FTIME		4	&& Original file time (hh:mm)
#define	BU_ANAME		5	&& Name of the backup file in
					&&    which the current file starts

*==========================================================================
*	bRBckBlk Return Codes
*==========================================================================
#define	BU_ABORT		0	&& Abort    Backup/Restore
#define	BU_CONT			1	&& Continue Backup/Restore

*==========================================================================
*	bConfirm Return Codes
*==========================================================================
#define	BU_SKIPFILE		2	&& Don't overwrite existing file
#define	BU_OVERWRITE		3	&& Overwrite existing file

*==========================================================================
*	Defines for the Progress Bar definition array (aBar)
*==========================================================================
#define	BAR_LEN			1	&& Length of the bar
#define	BAR_ROW			2	&& Row position
#define	BAR_COL			3	&& Column position
#define	BAR_CHR			4	&& Character to display
#define	BAR_CLR			5	&& Color

*==========================================================================
*	Defines for Insert Next Disk message definition array (aMsg)
*==========================================================================
#define	MSG_ROW			1	&& Row position
#define	MSG_COL			2	&& Column position
#define	MSG_TXT			3	&& Message text
#define	MSG_CLR			4	&& Color

*==========================================================================
* COMMAND FORM SYNTAX
*==========================================================================
#command RBACKUP [ <file:FILE,FILES> <aSkeleton> ]			  ;
		[ EXCLUDE <aExclude> ]					  ;
		[ <dest:TO,DEST,DESTINATION> <(cDestDrv)> ]		  ;
		[ <buset:BACKUPSET,FILENAME> <(cFileName)> ]		  ;
		[ BAR <aBar> ]						  ;
		[ <lEscape:ESCAPE,ESC,ESCAPE ON> ]			  ;
		[ ESCAPE OFF ]						  ;
		[ SILENT <lSilent> ]					  ;
		[ BLOCK <bRBckBlk> ]					  ;
		[ MESSAGE <aMsg> ]					  ;
		[ RETCODE <nRetCode> ]					  ;
		=>							  ;
	[ <nRetCode> := ] R_BackUp ( <aSkeleton>, <aExclude>,		  ;
		<(cDestDrv)>, <(cFileName)>,				  ;
		<aBar>, <.lEscape.>, <.lSilent.>, <bRBckBlk>, <aMsg> )

#command RRESTORE [ <src:FROM> <(cSrcDrv)> ]				  ;
		[ <file:FILE,FILES> <aMask> ]				  ;
		[ <dest:TO,DEST,DESTINATION> <(cDestSpec)> ]		  ;
		[ <buset:BACKUPSET,FILENAME> <(cFileName)> ]		  ;
		[ BAR <aBar> ]						  ;
		[ <lEscape:ESCAPE,ESC,ESCAPE ON> ]			  ;
		[ ESCAPE OFF ]						  ;
		[ SILENT <lSilent> ]					  ;
		[ BLOCK <bRBckBlk> ]					  ;
		[ MESSAGE <aMsg> ]					  ;
		[ <conf:CONFIRM,CONFIRMATION> <bConfirm> ]		  ;
		[ RETCODE <nRetCode> ]					  ;
		=>							  ;
	[ <nRetCode> := ] R_Restore ( <(cSrcDrv)>, <aMask>,		  ;
		<(cDestSpec)>, <(cFileName)>, <aBar>, <.lEscape.>,	  ;
		<.lSilent.>, <bRBckBlk>, <aMsg>, <bConfirm> )

#command RBCKLIST [ <src:FROM> <(cSrcDrv)> ]				  ;
		[ <buset:BACKUPSET,FILENAME> <(cFileName)> ]		  ;
		[ <lEscape:ESCAPE,ESC,ESCAPE ON> ]			  ;
		[ ESCAPE OFF ]						  ;
		[ SILENT <lSilent> ]					  ;
		[ MESSAGE <aMsg> ]					  ;
		[ RETCODE <nRetCode> ]					  ;
		=>							  ;
	[ <nRetCode> := ] R_BckList ( <(cSrcDrv)>, <(cFileName)>, 	  ;
		<.lEscape.>, <.lSilent.>, <aMsg> )
*
* Eof RBckLib.CH
