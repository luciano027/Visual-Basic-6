*****************************************************************************
* Programa...: RELPRC.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: VALORES - VISUALIZADO COM FONTE MAIOR - PARA VENDA A VISTA   *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************

function NGrande( nNGrande1, nNGrande2, nNGrande3, nNGrande4, cNGrande5 )

// nNGrande1 -> N£mero que dever† ser apresentado em tamanho grande
// nNGrande2 -> Linha Inicial para apresentaá∆o
// nNGrande3 -> Coluna Inicial para apresentaá∆o
// nNGrande4 -> Coluna Final para apresentaá∆o
// cNGrande5 -> M†scara para apresentaá∆o

cNGrande6 := transform( nNGrande1, cNGrande5 )

nNGrande7 := 1
nNGrande8 := len( cNGrande6 )

cNGrandeA := ""    //  Estas vari†veis conter∆o
cNGrandeB := ""    //  os caracteres que formar∆o
cNGrandeC := ""    //  os n£mero em tamanho
cNGrandeD := ""    //  grande.

do while nNGrande7 <= nNGrande8

    cNGrande9 := substr( cNGrande6, nNGrande7, 1 )

    do case
        case cNGrande9 == "0"

            cNGrandeA += "‹€€€€€‹ "
            cNGrandeB += "€€  ‹€€ "
            cNGrandeC += "€€‹ﬂ €€ "
            cNGrandeD += "ﬂ€€€€€ﬂ "

        case cNGrande9 == "1"

            cNGrandeA += "‹€€ "
            cNGrandeB += "ﬂ€€ "
            cNGrandeC += " €€ "
            cNGrandeD += " €€ "

        case cNGrande9 == "2"

            cNGrandeA += "‹€€€€€‹ "
            cNGrandeB += "  ‹‹€€ﬂ "
            cNGrandeC += "‹€€ﬂﬂ   "
            cNGrandeD += "€€€€€€€ "

        case cNGrande9 == "3"

            cNGrandeA += "€€€€€€‹ "
            cNGrandeB += "  ‹‹‹€€ "
            cNGrandeC += "  ﬂﬂﬂ€€ "
            cNGrandeD += "€€€€€€ﬂ "

        case cNGrande9 == "4"

            cNGrandeA += "€€  €€  "
            cNGrandeB += "€€‹‹€€‹ "
            cNGrandeC += "ﬂﬂﬂﬂ€€ﬂ "
            cNGrandeD += "    €€  "

        case cNGrande9 == "5"

            cNGrandeA += "€€€€€€€ "
            cNGrandeB += "€€‹‹‹‹  "
            cNGrandeC += " ﬂﬂﬂﬂ€€ "
            cNGrandeD += "€€€€€€ﬂ "

        case cNGrande9 == "6"

            cNGrandeA += "‹€€€€€‹ "
            cNGrandeB += "€€‹‹‹‹  "
            cNGrandeC += "€€ﬂﬂﬂ€€ "
            cNGrandeD += "ﬂ€€€€€ﬂ "

        case cNGrande9 == "7"

            cNGrandeA += "€€€€€€‹ "
            cNGrandeB += "   ‹€€ﬂ "
            cNGrandeC += "  €€ﬂ   "
            cNGrandeD += "  €€    "

        case cNGrande9 == "8"

            cNGrandeA += "‹€€€€€‹ "
            cNGrandeB += "ﬂ€‹‹‹€ﬂ "
            cNGrandeC += "€€ﬂﬂﬂ€€ "
            cNGrandeD += "ﬂ€€€€€ﬂ "

        case cNGrande9 == "9"

            cNGrandeA += "‹€€€€€‹ "
            cNGrandeB += "€€‹‹‹€€ "
            cNGrandeC += " ﬂﬂﬂ€€€ "
            cNGrandeD += " €€€€ﬂ  "

        case cNGrande9 == "."

            cNGrandeA += "   "
            cNGrandeB += "   "
            cNGrandeC += "   "
            cNGrandeD += "€€ "

        case cNGrande9 == ","

            cNGrandeA += "   "
            cNGrandeB += "   "
            cNGrandeC += "‹‹ "
            cNGrandeD += "€ﬂ "

    endcase

    nNGrande7++

enddo

nNGrandTam := ( nNGrande4 - nNGrande3 + 1 )

if len( cNGrandeA ) < nNGrandTam
    cNGrandeA := space( nNGrandTam - len( cNGrandeA ) ) + cNGrandeA
    cNGrandeB := space( nNGrandTam - len( cNGrandeB ) ) + cNGrandeB
    cNGrandeC := space( nNGrandTam - len( cNGrandeC ) ) + cNGrandeC
    cNGrandeD := space( nNGrandTam - len( cNGrandeD ) ) + cNGrandeD
endif

@ nNGrande2,nNGrande3 clear to nNGrande2+3,nNGrande4
set color to r/w
@ nNGrande2  , nNGrande3 say substr( cNGrandeA, 1, nNGrandTam )
@ nNGrande2+1, nNGrande3 say substr( cNGrandeB, 1, nNGrandTam )
@ nNGrande2+2, nNGrande3 say substr( cNGrandeC, 1, nNGrandTam )
@ nNGrande2+3, nNGrande3 say substr( cNGrandeD, 1, nNGrandTam )

return( nil )




