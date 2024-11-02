;*****************************************************
;*  EXAMPLE PROGRAM FOR ET-LAB3A                     *
;*  DESCRIPTION : READ & WRITE TO 24C16              *
;*  COMPILER    : SAX51                              *
;*****************************************************
            ORG     0000H
            LJMP    RESET_VEC               ;reset vector

SCL         EQU     P0.0
SDA         EQU     P0.1

CONT_BW     EQU     0A0H
CONT_BR     EQU     0A1H

            ORG     25H
ADDRESS:    DS      1
DATA:       DS      1

PORT_LCD    EQU     P1              ;Port  LCD
E_LCD       EQU     PORT_LCD.2      ;Enable LCD
WR_LCD      EQU     PORT_LCD.1      ;WR LCD
RS_LCD      EQU     PORT_LCD.0      ;RS Select LCD 



RESET_VEC:  MOV     SP,#256-32
            CLR     EA

INIT_PORT:  SETB    SCL
            SETB    SDA

            CLR     WR_LCD          ;write LCD
            SETB    E_LCD           ;initial LCD

            LCALL   DELAY           ;initial delay

            LCALL   INIT_LCD        ;initial LCD

            MOV     DPTR,#TEXT
SSP2:       MOV     A,#0
            MOVC    A,@A+DPTR
            CJNE    A,#0DH,SSP
            LJMP    SSP1
SSP:        LCALL   WRITE_DATA
            INC     DPTR
            LJMP    SSP2

SSP1:       MOV     A,#40H
            LCALL   GOTOLCD
            MOV     A,#35H
            LCALL   WRITE_DATA
            MOV     A,#30H
            LCALL   WRITE_DATA


            MOV     ADDRESS,#0
            MOV     DATA,#50H
            LCALL   WRITE_BYTE

LOOP:       MOV     ADDRESS,#00H
            LCALL   READ_BYTE

            MOV     A,#49H
            LCALL   GOTOLCD

            MOV     A,DATA
            SWAP    A
            ANL     A,#0FH
            LCALL   HEXASC
            LCALL   WRITE_DATA

            MOV     A,DATA
            ANL     A,#0FH
            LCALL   HEXASC
            LCALL   WRITE_DATA

            SJMP    LOOP


;*****************************************************
;*            WRITE DATA TO EEPROM 1 BYTE            *
;*      : DATA                                       *
;*****************************************************
WRITE_BYTE: CLR     SDA                    ;start bit
            CLR     SCL

            MOV     A,#CONT_BW             ;send control byte
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,WRITE_BYTE         ;loop until busy
            CLR     SCL
            NOP
            NOP

            MOV     A,ADDRESS              ;send address low
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,WRITE_BYTE         ;loop until busy
            CLR     SCL
            NOP
            NOP

            MOV     A,DATA                 ;send data
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,WRITE_BYTE         ;loop until busy
            CLR     SCL
            NOP
            NOP
            CLR     SDA

            SETB    SCL                    ;stop bit
            SETB    SDA
            LCALL   DELAY
            RET

;******************************************************
;*            READ DATA FROM EEPROM 1 BYTE            *
;* OUTPUT : DATA                                      *
;******************************************************
READ_BYTE:  CLR     SDA                    ;start bit
            CLR     SCL

            MOV     A,#CONT_BW             ;send control byte
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,READ_BYTE          ;loop until busy
            CLR     SCL
            NOP
            NOP

            MOV     A,ADDRESS              ;send address low
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,READ_BYTE         ;loop until busy
            CLR     SCL
            NOP
            NOP

            SETB    SCL

            CLR     SDA                    ;start bit
            CLR     SCL

            MOV     A,#CONT_BR             ;send control byte
            LCALL   PUT_BYTE
            SETB    SDA
            SETB    SCL
            NOP
            NOP
            JB      SDA,READ_BYTE          ;loop until busy
            CLR     SCL
            NOP
            NOP

            LCALL   GET_BYTE
            MOV     DATA,A

            SETB    SDA
            SETB    SCL                    ;NO ACK
            CLR     SCL

            CLR     SDA

            SETB    SCL                   ;stop bit
            SETB    SDA
            LCALL   DELAY
            RET


;****************************************************
;*                      WRITE                       *
;* INPUT: ACC                                       *
;****************************************************
PUT_BYTE:   PUSH    02H
            MOV     R2,#08H
PUT1:       RLC     A
            MOV     SDA,C
            NOP
            NOP
            SETB    SCL
            NOP
            NOP
            CLR     SCL
            NOP
            NOP
            DJNZ    R2,PUT1
            POP     02H
            RET

;*****************************************************
;*                       READ                        *
;* OUTPUT: ACC                                       *
;*****************************************************
GET_BYTE:   PUSH    02H
            MOV     R2,#08H
GET1:       SETB    SCL
            NOP
            NOP
            MOV     C,SDA
            CLR     SCL
            NOP
            NOP
            RLC     A
            DJNZ    R2,GET1
            POP     02H
            RET

;***************************************************
;*                     delay
;***************************************************
DELAY4M:    PUSH    DPH
            PUSH    DPL
            MOV     DPTR,#0000H
DEL4:       INC     DPTR
            MOV     A,DPL
            ORL     A,DPH
            JNZ     DEL4
            POP     DPL
            POP     DPH
            RET


;****************************************************
;*                Write Instruction LCD             *
;*                Input  : ACC (Command)            *
;*                Output : Data bus LCD             *
;****************************************************
WRITE_INS:  MOV     B,A
            ANL     A,#0F0H
            MOV     P1,A
            LCALL   EN_LCD
            MOV     A,B
            SWAP    A
            ANL     A,#0F0H
            MOV     P1,A
            LCALL   EN_LCD
            RET

WRITE_DATA: MOV     B,A
            ANL     A,#0F0H
            SETB    ACC.0
            MOV     P1,A
            LCALL   EN_LCD
            MOV     A,B
            SWAP    A
            ANL     A,#0F0H
            SETB    ACC.0
            MOV     P1,A
            LCALL   EN_LCD
            RET 

;****************************************************
;*                      Initial LCD                 *
;*                      4-Bit Interface             *
;****************************************************
INIT_LCD:   CLR     RS_LCD
            MOV     A,#33H       
            LCALL   WRITE_INS
            MOV     A,#32H       
            LCALL   WRITE_INS
            MOV     A,#28H        
            LCALL   WRITE_INS         
            MOV     A,#01H
            LCALL   WRITE_INS         
            MOV     A,#02H          
            LCALL   WRITE_INS          
            MOV     A,#0FH         
            LCALL   WRITE_INS          
            RET


GOTOLCD:    SETB   ACC.7
            LCALL  WRITE_INS
            RET

;****************************************************
;*                    Enable Pin E LCD              *
;*                    Active Chip select            *
;****************************************************
EN_LCD:     SETB    E_LCD           ;Enable LCD
            LCALL   BUSY            ;Busy delay time
            CLR     E_LCD           ;Disable LCD
            LCALL   BUSY
            RET

;****************************************************
;*                Delay time for Busy               * 
;*                  Wait LCD Ready                  *
;****************************************************
BUSY:       PUSH    07H
            PUSH    06H
            MOV     R6,#05H
BUSY1:      MOV     R7,#0FFH
            DJNZ    R7,$
            DJNZ    R6,BUSY1
            POP     06H
            POP     07H
            RET


;****************************************************
;*                      Delay Time                  *
;****************************************************
DELAY:      PUSH    DPH
            PUSH    DPL
            PUSH    ACC
            MOV     DPTR,#0000H
DEL:        INC     DPTR
            MOV     A,DPL
            ORL     A,DPH
            JNZ     DEL
            POP     ACC
            POP     DPL
            POP     DPH
            RET

HEXASC:     CJNE    A,#10,CHECK
CHECK:      JNC     MORE
            ADD     A,#30H
            RET
MORE:       ADD     A,#37H 
            RET

TEXT:       DB      'WRITE   READ',0DH

            END














