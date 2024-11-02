;********************************************
;*  EXAMPLE PROGRAM FOR ET-LAB3A            *
;*  DESCRIPTION : SCAN 7-SEGMENT            *
;*  COMPILER    : SXA51                     *
;********************************************
              ORG   000H
              LJMP  MAIN

MAIN:         MOV   SP,#256-32               ;define stack 32 byte
              CLR   EA

LOOP1:        MOV   DPTR,#TAB
              MOV   R2,#5
              MOV   R1,#0

LOOP:         MOV   P0,#00H
              MOV   P1,R1
              LCALL DELAY 
              CLR   A
              MOVC  A,@A+DPTR
              MOV   P0,A
              MOV   P1,R1
              LCALL DELAY
              INC   DPTR
              INC   R1
              DJNZ  R2,LOOP
              LJMP  LOOP1

;*****************************************
;*               delay                   *
;*****************************************
DELAY:        PUSH  DPH
              PUSH  DPL
              PUSH  ACC
              MOV   DPTR,#0FFF0H
DEL:          INC   DPTR
              MOV   A,DPL
              ORL   A,DPH
              JNZ   DEL
              POP   ACC
              POP   DPL
              POP   DPH
              RET


TAB:          DB    06H,5BH,4FH,66H,00H          ;table

              END

