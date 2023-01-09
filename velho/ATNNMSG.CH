*********** Sample programs for CA-Clipper Tools / 5.2 ************
*
* File	      : CTNNMSG.CH
* Copyright   : (c) 1992-93, Computer Associates International.
*		All rights reserved.
* Date	      : 05/20/92
* Author      : Michael Peters
* Description : Header file for MAP.PRF
*************************************************************************

#ifndef _NTNNMSG_DEFINED

#define ERR_MSG_01 "You have reached the maximum number of server attachements."
#define ERR_MSG_02 "You are already attached to server %1 as %2."
#define ERR_MSG_03 "Intruder detection lockout has disabled this account."
#define ERR_MSG_04 "Password for user %1 on server %2 has expired."
#define ERR_MSG_05 "   You have NO grace logins left."
#define ERR_MSG_06 "   This is you LAST CHANCE to change your password."
#define ERR_MSG_07 "Control characters not allowed in command."
#define ERR_MSG_08 "Cannot get connection status (%1)."
#define ERR_MSG_09 "Unable to logout from %1 as %2 (%3)."
#define ERR_MSG_10 "%1 logged out from server %2"
#define ERR_MSG_11 "Access denied."
#define ERR_MSG_12 "Attempting to login from an unapproved station." +CHR(13) +CHR(10) +;
		   " The supervisor has limited the stations you are allowed to login on."
#define ERR_MSG_13 "Attempting to login during an unauthorized time period."  +CHR(13) +CHR(10) +;
		   " The supervisor has limited the times you can login to this server."
#define ERR_MSG_14 "Attempting to simultaneously login to too many work stations."  +CHR(13) +CHR(10) +;
		   " The supervisor has limited the number of active connections you may have."
#define ERR_MSG_15 "Password has expired and all grace logins have been used."
#define ERR_MSG_16 "   You have %1 grace login(s) left to change your password."
#define ERR_MSG_17 "This account has expired or been disabled by the supervisor."
#define ERR_MSG_18 "You have exceeded your credit limit for this server."
#define ERR_MSG_19 "You do not have an account balance. This server charges for file services."
#define ERR_MSG_20 "The file server has reached its maximum number of connections."
#define ERR_MSG_21 "Unknown file server."
#define ERR_MSG_22 "No response from file server."
#define ERR_MSG_23 "The supervisor has disabled the login function for this server."
#define ERR_MSG_24 "New password not retyped correctly."
#define ERR_MSG_25 "New password is too short."
#define ERR_MSG_26 "The new password has been used previously."
#define ERR_MSG_27 "Unknown error changing password (%1)."
#define ERR_MSG_28 "Unable to detach from file server (%1)."
#define ERR_MSG_29 "Unable to get connection list (%1)."
#define ERR_MSG_30 "No user %1 on file server %2"
#define ERR_MSG_31 "You are attached to the maximum number of servers allowed."  +CHR(13) +CHR(10) +;
		   "You must logout from a server before attaching to another server."
#define ERR_MSG_32 "Unknown error during attachment (%1)."
#define ERR_MSG_33 "%1: Access to server denied and you have been logged out."


#define MSG_01	   "Would you like to change you password on server %1 (Y/N) "
#define MSG_02	   "You are attached to server %1"
#define MSG_03	   "Username: "
#define MSG_04	   "Server:   "
#define MSG_05	   "Password: "
#define MSG_06	   "Do you still want to attach as %1 (Y/N) "
#define MSG_07	   "YES"
#define MSG_08	   "NO"
#define MSG_09	   "Copyright (c) 1899, Yesvell, Inc. All rights reversed"
#define MSG_10	   "Enter your new password:  "
#define MSG_11	   "Retype your new password: "
#define MSG_12	   "Synchronize passwords with %1/%2 on these file servers? (Y/N) "
#define MSG_13	   "Your password has been changed."
#define MSG_14	   "Would you still like to change your password on server %1? (Y/N) "
#define MSG_15	   "Enter your login name:"
#define MSG_16	   "Enter your password:"

#define _NTNNMSG_DEFINED
#endif