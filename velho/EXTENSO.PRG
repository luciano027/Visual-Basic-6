*********************************************
* Funcao que imprime numero por extenso
* Luis Aguiar
*********************************************
do while .t.
   kvalor=0
   cls
   @ 1,1 to 24,79 double
   @ 8,10 say'ENTRE COM O VALOR' GET kvalor pict '999,999,999.99'
   READ
   if lastkey()=27
      cls
      quit
   endif
   convext(kvalor,10,2)
   @ 23,10 say 'Enter para continuar...'
   inkey(0)
enddo



*************************************
FUNCTION CONVEXT
PARAMETERS kvalor,lin,col
public kvalor,lin,col
EXT = "("
A1="Cem"
A2="Duzentos"
A3="Trezentos"
A4="Quatrocentos"
A5="Quinhentos"
A6="Seiscentos"
A7="Setecentos"
A8="Oitocentos"
A9="Novecentos"
B1="Um"
B2="Dois"
B3="Tres"
B4="Quatro"
B5="Cinco"
B6="Seis"
B7="Sete"
B8="Oito"
B9="Nove"
B10="Dez"
B11="Onze"
B12="Doze"
B13="Treze"
B14="Quatorze"
B15="Quinze"
B16="Dezesseis"
B17="Dezessete"
B18="Dezoito"
B19="Dezenove"
C2="Vinte"
C3="Trinta"
C4="Quarenta"
C5="Cinquenta"
C6="Sessenta"
C7="Setenta"
C8="Oitenta"
C9="Noventa"
P=1
STG=0
V1=SUBSTR(STR(kvalor,12,2),1,3)
V2=SUBSTR(STR(kvalor,12,2),4,3)
V3=SUBSTR(STR(kvalor,12,2),7,3)
V4=" "+SUBSTR(STR(kvalor,12,2),11,2)
DO WHILE P<5


   ST1=0
   ST2=0
   ST3=0
   PA="V"+STR(P,1,0)
   VA=&PA
   CT=1
   DO WHILE VAL(SUBSTR(VA,1,1))#0
      IF VAL(SUBSTR(VA,1,1))=CT
         IF STG=1
            IF P#4
               CONEC=" e "
            ENDIF
         ENDIF
         EXT1="A"+STR(CT,1,0)
         EXT1=&EXT1
         IF VAL(SUBSTR(VA,1,1))=1 .AND. VAL(SUBSTR(VA,2,2))#0
            EXT1="Cento"
         ENDIF
         IF STG=0
            EXT=EXT+EXT1
         ELSE
            EXT=EXT+CONEC+EXT1
         ENDIF
         ST1=1
         EXIT
      ENDIF
      CT=CT+1
      IF CT=10
         EXIT
      ENDIF
   ENDDO
   CT=2
   DO WHILE VAL(SUBSTR(VA,2,1))#0 .AND. VAL(SUBSTR(VA,2,1))#1
      IF VAL(SUBSTR(VA,2,1))=CT
         EXT1="C"+STR(CT,1,0)
         IF ST1=0 .AND. STG=0
            EXT=EXT+&EXT1
         ELSE
            EXT=EXT+" e "+&EXT1
         ENDIF
         ST2=1
         EXIT
      ENDIF
      CT=CT+1
      IF CT=10
         EXIT
      ENDIF
   ENDDO
   CT=10
   DO WHILE VAL(SUBSTR(VA,2,1))=1
      IF VAL(SUBSTR(VA,2,2))=CT
         EXT1="B"+LTRIM(STR(CT,2,0))
         IF ST1=0 .AND. ST2=0 .AND. STG=0
            EXT=EXT + &ext1
         ELSE
            EXT=EXT+" e " + &EXT1
         ENDIF
         ST3=1
         EXIT
      ENDIF
      CT=CT+1
      IF CT=20
         EXIT
      ENDIF
   ENDDO
   CT=1
   DO WHILE VAL(SUBSTR(VA,2,1))=0 .OR. VAL(SUBSTR(VA,2,1))#1
      IF VAL(SUBSTR(VA,3,1))=CT
         EXT1="B" + LTRIM(STR(CT,2,0))
         IF ST1=0 .AND. ST2=0 .AND. STG=0
            EXT=EXT+&EXT1
         ELSE
            EXT=EXT+" e " +&EXT1
         ENDIF
         ST3=1
         EXIT
      ENDIF
      CT=CT+1
      IF CT=10
         EXIT
      ENDIF
   ENDDO
   CT=1
   IF ST1=1 .OR. ST2=1 .OR. ST3=1 .OR. (P=3 .AND. INT(kvalor)#0)
      IF P=1
         IF VAL(VA)=1
            EXT=EXT+" Milhao"
         ELSE
            EXT=EXT+" Milhoes"
         ENDIF
      ENDIF
      IF P=2
         EXT=EXT+" Mil"
      ENDIF
      IF P=3
         IF INT(kvalor)=1
            EXT=EXT+" Real"
         ELSE
            EXT=EXT+" Reais"
         ENDIF
      ENDIF
      IF P=4
         IF VAL(VA)=1
            EXT=EXT+" Centavo"
         else
            EXT=EXT+" Centavos"
         ENDIF
      ENDIF
      STG=1
   ENDIF
   P=P+1
ENDDO
EXT=EXT+")"
@ lin,col say ext
RETURN //(EXT)
