;*************************************
;* EXAMPLE PROGRAM FOR ET-LAB3A      *
;* DESCRIPTION : SEND DATA TO DAC    *
;* COMPILER    : SAX51               *
;*************************************
           ORG    0000H   
           LJMP   MAIN

MAIN:      MOV    SP,#256-32         ;define stack 32 byte 

LOOP:      MOV    P0,#00H
           LCALL  DELAY
           MOV    P0,#80H
           LCALL  DELAY
           MOV    P0,#0FFH
           LCALL  DELAY

           LJMP   LOOP

;**************************************
;*               delay                *
;**************************************
DELAY:     MOV    R2,#10
DEL1:      MOV    DPTR,#0000H
DEL:       INC    DPTR
           MOV    A,DPH
           ORL    A,DPL
           JNZ    DEL
           DJNZ   R2,DEL1
           RET

           END
