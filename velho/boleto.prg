#include "common.ch"
#include "inkey.ch"

********************************
function EMIS_BOL

   private mes:= Space(7), cd_in:= cd_fin:= Space(8)
   if (abrearq("C_COND02", "COND", .F., 10))
      set index to GEC02_01, GEC02_02
   else
      mensagem("O arquivo nao esta disponivel", 3, 2)
      close databases
      select 1
      close format
      return
   endif
   if (abrearq("C_DEB03", "DEBITOS", .F., 10))
      set index to GEC03_01
   else
      mensagem("O arquivo nao esta disponivel", 3, 2)
      close databases
      select 1
      close format
      return
   endif
   if (abrearq("C_LANC04", "LANC", .F., 10))
      set index to GEC04_01
   else
      mensagem("O arquivo nao esta disponivel", 3, 2)
      close databases
      select 1
      close format
      return
   endif
   if (abrearq("C_DLAN05", "DEB_LANC", .F., 10))
      set index to GEC05_01, GEC05_02
   else
      mensagem("O arquivo nao esta disponivel", 3, 2)
      close databases
      select 1
      close format
      return
   endif
   if (abrearq("C_LCON06", "COND_LANC", .F., 10))
      set index to GEC06_01, GEC06_02
   else
      mensagem("O arquivo nao esta disponivel", 3, 2)
      close databases
      select 1
      close format
      return
   endif
   do while (.T.)
      abrejan(2)
      set color to (vcr3)
      clear screen
      readkill(.T.)
      getlist:= {}
      cab("Emissao de Boletos")
      set color to (vcr3)
      @  8, 28 say "Formul rio Especial"
      @ 20,  5 say ;
         "Obs.: Para emissao de boletos no formulario, NAO imprimir via VIDEO"
      @ 10,  8 to 18, 75
      @ 12, 10 say "Periodo________: " get mes picture "99/9999"
      @ 14, 10 say "Codigo Inicial_: " get cd_in picture "999!/999" ;
         valid cdin()
      @ 16, 10 say "Codigo Final___: " get cd_fin picture "999!/999" ;
         valid cdfin()
      le()
      if (LastKey() == K_ESC)
         close databases
         select 1
         close format
         return
      endif
      select COND_LANC
      set order to 2
      seek mes + cd_in softseek
      if (EOF() .OR. a06_mes + a06_codc > mes + cd_fin)
         mensagem("Intervalo inexistente, redigite", 2, 1)
         loop
      endif
      exit
   enddo
   set color to (vcr)
   caixa(0, 5, 4, 21, frame[2])
   @  1,  7 prompt "1- Video      "
   @  2,  7 prompt "2- Impressora "
   @  3,  7 prompt "3- Menu       "
   menu to video
   if (video == 1)
      tipo_prn:= "T"
      ex_t:= Val(SubStr(Time(), 4, 2)) * 10 + Val(SubStr(Time(), 7, ;
         2))
      arq_prn:= "list" + alltrim(Str(ex_t))
      set printer to (arq_prn)
      caixa(12, 20, 14, 55, frame[2])
      @ 13, 22 say "Aguarde, processando relatorio"
      set color to (vcr3)
   elseif (video == 2)
      set color to (vcr3)
      clear screen
      readkill(.T.)
      getlist:= {}
      if (!imprime("Emissao de Boletos"))
         close databases
         select 1
         close format
         return
      endif
      tipo_prn:= "I"
   else
      set filter to
      close databases
      select 1
      close format
      return
   endif
   set device to printer
   pg:= 0
   @ PRow(),  0 say inic_imp
   @ PRow(),  0 say ""
   do while (!EOF() .AND. a06_mes + a06_codc >= mes + cd_in .AND. ;
         a06_mes + a06_codc <= mes + cd_fin)
      if (escprint())
         close databases
         select 1
         close format
         set device to printer
         @ PRow(),  0 say ""
         set device to screen
         return Nil
      endif
      select COND
      set order to 1
      seek cond_lanc->a06_codc
      select LANC
      set order to 1
      seek cond_lanc->a06_codl + cond_lanc->a06_mes
      select COND_LANC
      @ PRow() + 1, 51 say lanc->a04_dtven
      @ PRow(),  0 say red_1
      @ PRow() + 4,  3 say lanc->a04_dtdoc
      @ PRow(), PCol() + 9 say a06_codc
      @ PRow(), PCol() + 10 say "REC."
      @ PRow(), PCol() + 22 say right(DToC(lanc->a04_dtven), 7)
      @ PRow(),  0 say tir_r1
      @ PRow() + 1, 54 say lanc->a04_valor picture "@E 9,999.99"
      @ PRow() + 1,  0 say red_1
      select DEB_LANC
      set order to 1
      seek cond_lanc->a06_codl + cond_lanc->a06_mes
      do while (!EOF() .AND. a05_codl + a05_mes = ;
            cond_lanc->a06_codl + cond_lanc->a06_mes)
         select DEBITOS
         set order to 1
         seek deb_lanc->a05_codd
         select DEB_LANC
         @ PRow() + 1,  3 say debitos->a03_desc
         @ PRow(), PCol() + 1 say a05_valor picture "@E 999.99"
         skip 
         if (!EOF() .AND. a05_codl + a05_mes = cond_lanc->a06_codl + ;
               cond_lanc->a06_mes)
            select DEBITOS
            set order to 1
            seek deb_lanc->a05_codd
            select DEB_LANC
            @ PRow(), PCol() + 3 say debitos->a03_desc
            @ PRow(), PCol() + 1 say a05_valor picture "@E 999.99"
            skip 
         endif
      enddo
      @ PRow() + 1,  3 say lanc->a04_obs1
      @ PRow() + 1,  3 say lanc->a04_obs2
      @ 15, 15 say cond->a02_nom
      @ 16, 15 say cond->a02_end
      @ 16, PCol() + 1 say alltrim(cond->a02_cid) + "  -"
      @ 16, PCol() + 1 say cond->a02_est
      @ PRow() + 8,  0 say ""
      @ PRow(),  0 say tir_r1
      select COND_LANC
      setprc(0, 0)
      skip 
   enddo
   eject
   @ PRow() + 1,  0 say " "
   @ PRow(),  0 say inic_imp
   setprc(0, 0)
   set printer to 
   set device to screen
   close databases
   select 1
   close format
   if (tipo_prn = "T")
      imp_tela(arq_prn, 80)
   endif
   return Nil

* EOF
