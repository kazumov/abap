*
* glu1 - https://www.sapdatasheet.org/abap/tabl/glu1.html (G/L user table 1)
* bkpf - https://www.sapdatasheet.org/abap/tabl/bkpf.html (Accounting Document Header)
* BSEG - https://www.sapdatasheet.org/abap/tabl/bseg.html (Accounting Document Segment)
*

DATA:
   ld_i_docnr  	TYPE GLU1-DOCNR,
   it_t_glu1  	TYPE STANDARD TABLE OF GLU1,"tables param
   wa_t_glu1	LIKE LINE OF it_t_glu1 .



DATA(ld_e_docnr) = '0123456789' "GLU1-DOCNR. Type CHAR, length 10.
DATA(ld_e_bukrs) = 'abcd'       "GLU1-BUKRS. Type CHAR, length 4.
DATA(ld_e_year) = '2018'        "GLU1-RYEAR. Type NUMC, length 4.

SELECT single BUDAT 
  FROM BKPF 
  INTO @DATA(ld_e_budat).


SELECT single BUDAT 
  FROM BKPF 
  INTO @DATA(ld_e_wudat).


SELECT single BUZID 
  FROM BSEG 
  INTO @DATA(ld_e_zerochk).
 

"populate fields of struture and append to itab
append wa_t_glu1 to it_t_glu1.

CALL FUNCTION 'JV_GB01_DIRECT_INPUT'
  EXPORTING
*   e_docnr =                    "ld_e_docnr - Accounting Document Number: GLU1-DOCNR
    e_bukrs = ld_e_bukrs         "- Company Code: GLU1-BUKRS 
    e_year = ld_e_year           "- Fiscal Year: BKPF-GJAHR, GLU1-RYEAR
*   e_budat = sy-datum           "ld_e_budat - Posting Date in the Document: BKPF-BUDAT, sy-satum
*   e_wudat = 00000000           "ld_e_wudat - Posting Date in the Document: BKPF-BUDAT, 00000000
*   e_zerochk =                  "ld_e_zerochk - Identification of the Line Item: BSEG-BUZID
  IMPORTING
    i_docnr = ld_i_docnr         "ld_e_docnr - Accounting Document Number: GLU1-DOCNR. CHAR(1)
  TABLES
    t_glu1 = it_t_glu1
  EXCEPTIONS
    DOCUMENT_NUMBER_NOT_FOUND =    1
    LOCAL_CURRENCY_NOT_CORRECT =   2
    .  "  JV_GB01_DIRECT_INPUT

IF SY-SUBRC EQ 0.
  "All OK
ELSEIF SY-SUBRC EQ 1. "Exception
  "Add code for exception here
ELSEIF SY-SUBRC EQ 2. "Exception
  "Add code for exception here
ENDIF.
