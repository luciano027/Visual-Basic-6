/*
    File: GRUMPM.CH
    System: GRUMPFISH LIBRARY
    Author: Greg Lief
    Copyright (c) 1988-90, Greg Lief
    Header file to be used in conjunction with Grumpfish Library's LITE_MENU2()
    This provides new definitions for Clipper's @..PROMPT and MENU TO commands.
*/

//ÄÄÄÄÄ optional DISPLAY keyword will cause option to be displayed
//ÄÄÄÄÄ immediately rather than waiting until inside Lite_Menu2()
#xcommand @ <row>, <col> PROMPT <prompt> [MESSAGE <msg>] [WHEN <when>] ;
           [MESSAGECOLOR <msgcolor>] [ACTION <action> [TO <var>]]      ;
           [TRIGGERCOLOR <trigcolor>] [DISPLAY]                        ;
                                                                   =>  ;
           if(menulist == NIL, menulist := {}, )                     ; ;
           setpos(<row>, <col>)                                      ; ;
           aadd(menulist, { row(), col(), <prompt>, <msg>, <msgcolor>, ;
                           <{when}>, <{action}>, <trigcolor>,          ;
                           [ { | _1 | <var> := _1 } ] } )            ; ;
           dispout( <prompt> )

#xcommand @ <row>, <col> PROMPT <prompt> [MESSAGE <msg>] [WHEN <when>] ;
           [MESSAGECOLOR <msgcolor>] [ACTION <action> [TO <var>]]      ;
           [TRIGGERCOLOR <trigcolor>]                                  ;
                                                                   =>  ;
           if(menulist == NIL, menulist := {}, )                     ; ;
           setpos(<row>, <col>)                                      ; ;
           aadd(menulist, { row(), col(), <prompt>, <msg>, <msgcolor>, ;
                           <{when}>, <{action}>, <trigcolor>,          ;
                           [ { | _1 | <var> := _1 } ] } )


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [CLOCK] [TIMEOUT <timeout>]                  ;
                      [PULLDOWN]                                   ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>] [SAVE]            ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                        {|| gfclock()}, <timeout>, <{exitevent}>,  ;
                        <trigcolor>, chr(27)+chr(19)+chr(13),      ;
                                     chr(27)+chr(4)+chr(13),       ;
                        ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>)


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [CLOCK] [TIMEOUT <timeout>]                  ;
                      [LEFT <left>] [RIGHT <right>]                ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>] [SAVE]            ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                        {|| gfclock()}, <timeout>, <{exitevent}>,  ;
                        <trigcolor>, <left>, <right>,              ;
                        ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>)


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [EVENT <event>] [TIMEOUT <timeout>]          ;
                      [PULLDOWN]                                   ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>] [SAVE]            ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                           <{event}>, <timeout>, <{exitevent}>,    ;
                           <trigcolor>, chr(27)+chr(19)+chr(13),   ;
                           chr(27)+chr(4)+chr(13), ! <.center.>,   ;
                           ! <.clrmsg.>, <mrow>, <mcol>)


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [EVENT <event>] [TIMEOUT <timeout>]          ;
                      [LEFT <left>] [RIGHT <right>]                ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>] [SAVE]            ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                        <{event}>, <timeout>, <{exitevent}>,       ;
                        <trigcolor>, <left>, <right>,              ;
                        ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>)


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [CLOCK] [TIMEOUT <timeout>]                  ;
                      [PULLDOWN]                                   ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>]                   ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                        {|| gfclock()}, <timeout>, <{exitevent}>,  ;
                        <trigcolor>, chr(27)+chr(19)+chr(13),      ;
                                     chr(27)+chr(4)+chr(13),       ;
                        ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>);
         ; menulist := {}


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [CLOCK] [TIMEOUT <timeout>]                  ;
                      [LEFT <left>] [RIGHT <right>]                ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>]                   ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                      {|| gfclock()}, <timeout>, <{exitevent}>,    ;
                      <trigcolor>, <left>, <right>,                ;
                      ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>)  ;
         ; menulist := {}


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [EVENT <event>] [TIMEOUT <timeout>]          ;
                      [PULLDOWN]                                   ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>]                   ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                           <{event}>, <timeout>, <{exitevent}>,    ;
                           <trigcolor>, chr(27)+chr(19)+chr(13),   ;
                           chr(27)+chr(4)+chr(13), ! <.center.>,   ;
                           ! <.clrmsg.>, <mrow>, <mcol>)           ;
         ; menulist := {}


#xcommand MENU TO <v> [COLOR <color>] [INITIAL <initial>]          ;
                      [EVENT <event>] [TIMEOUT <timeout>]          ;
                      [LEFT <left>] [RIGHT <right>]                ;
                      [EXITEVENT <exitevent>]                      ;
                      [TRIGGERCOLOR <trigcolor>]                   ;
                      [<center:NOCENTER_MESSAGES>]                 ;
                      [<clrmsg:NOCLEAR_MESSAGES>]                  ;
                      [MESSAGEROW <mrow>]                          ;
                      [MESSAGECOL <mcol>]                          ;
                                                                => ;
         <v> := lite_menu2(menulist, <color>, <initial>, #<v>,     ;
                       <{event}>, <timeout>, <{exitevent}>,        ;
                       <trigcolor>, <left>, <right>,               ;
                       ! <.center.>, ! <.clrmsg.>, <mrow>, <mcol>) ;
         ; menulist := {}


* eof: GRUMPM.CH
