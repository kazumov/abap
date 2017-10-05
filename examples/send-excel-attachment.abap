*&---------------------------------------------------------------------*
*& Report ZSDRR_DAILY_FEED
*&
*&---------------------------------------------------------------------*
*& Description : This is the report program to send email with *
*& list of new contract created daily, weekly and *
*& monthly to specified email ID in the selection *
*& screen. *
*&---------------------------------------------------------------------*
REPORT zsdrr_daily_feed.
TABLES: mara.
*/..Types Declarations
TYPES : BEGIN OF ty_vbak,
vbeln TYPE vbak-vbeln,
erdat TYPE vbak-erdat,
auart TYPE vbak-auart,
augru TYPE vbak-augru,
vkbur TYPE vbak-vkbur,
kunnr TYPE vbak-kunnr,
END OF ty_vbak,
BEGIN OF ty_vbkd,
vbeln TYPE vbkd-vbeln,
posnr TYPE vbkd-posnr,
zzreinstate TYPE vbkd-zzreinstate,
zzexpired TYPE vbkd-zzexpired,
zzlatesale TYPE vbkd-zzlatesale,
zzrenewal TYPE vbkd-zzrenewal,
END OF ty_vbkd,
BEGIN OF ty_veda,
vbeln TYPE veda-vbeln,
vposn TYPE veda-vposn,
vbegdat TYPE veda-vbegdat,
venddat TYPE veda-venddat,
vkuegru TYPE veda-vkuegru,
vbedkue TYPE veda-vbedkue,
END OF ty_veda,
BEGIN OF ty_vbap,
vbeln TYPE vbap-vbeln,
posnr TYPE vbap-posnr,
matnr TYPE vbap-matnr,
zmeng TYPE vbap-zmeng,
kzwi3 TYPE vbap-kzwi3,
kzwi5 TYPE vbap-kzwi5,
maktx TYPE makt-maktx,
mtart TYPE mara-mtart,
END OF ty_vbap,
BEGIN OF ty_vbpa,
vbeln TYPE vbpa-vbeln,
posnr TYPE vbpa-posnr,
parvw TYPE vbpa-parvw,
kunnr TYPE vbpa-kunnr,
pernr TYPE vbpa-pernr,
adrnr TYPE vbpa-adrnr,
END OF ty_vbpa,
BEGIN OF ty_vbpa_s,
sobid TYPE hrp1001-sobid,
kunnr TYPE but000-partner,
END OF ty_vbpa_s,
BEGIN OF ty_adrc,
addrnumber TYPE adrc-addrnumber,
name1 TYPE adrc-name1,
name2 TYPE adrc-name2,
street TYPE adrc-street,
city1 TYPE adrc-city1,
post_code1 TYPE adrc-post_code1,
country TYPE adrc-country,
region TYPE adrc-region,
tel_number TYPE adrc-tel_number,
END OF ty_adrc,
BEGIN OF ty_adr6,
addrnumber TYPE adrc-addrnumber,
smtp_addr TYPE adr6-smtp_addr,
END OF ty_adr6,
BEGIN OF ty_but000,
partner TYPE but000-partner,
bpkind TYPE but000-bpkind,
bu_group TYPE but000-bu_group,
name_last TYPE but000-name_last,
name_first TYPE but000-name_first,
END OF ty_but000,
BEGIN OF ty_hrp1001,
otype TYPE hrp1001-otype,
objid TYPE hrp1001-objid,
plvar TYPE hrp1001-plvar,
relat TYPE hrp1001-relat,
begda TYPE hrp1001-begda,
endda TYPE hrp1001-endda,
sclas TYPE hrp1001-sclas,
sobid TYPE hrp1001-sobid,
END OF ty_hrp1001,
BEGIN OF ty_vbfa,
vbelv TYPE vbfa-vbelv,
posnv TYPE vbfa-posnv,
vbeln TYPE vbfa-vbeln,
posnn TYPE vbfa-posnn,
vbtyp_n TYPE vbfa-vbtyp_n,
vbtyp_v TYPE vbfa-vbtyp_v,
END OF ty_vbfa,
BEGIN OF ty_vbeln,
vbeln TYPE vbak-vbeln,
posnr TYPE vbap-posnr,
END OF ty_vbeln.
TYPES: BEGIN OF ty_kna1,
kunnr TYPE kna1-kunnr,
erdat TYPE kna1-erdat,
END OF ty_kna1.
TYPES: BEGIN OF x_final,
auart TYPE vbak-auart,
erdat TYPE vbak-erdat,
augru TYPE vbak-augru,
vbeln TYPE vbak-vbeln,
posnr TYPE vbap-posnr,
kunnr TYPE vbak-kunnr,
name(60) TYPE c,
telf1 TYPE sza1_d0100-tel_number,
smtp_addr TYPE sza1_d0100-smtp_addr,
street TYPE addr1_data-street,
city1 TYPE addr1_data-city1,
region TYPE addr1_data-region,
post_code1 TYPE addr1_data-post_code1,
country TYPE addr1_data-country,
kunnr_b TYPE vbak-kunnr,
name_b(60) TYPE c,
telf1_b TYPE sza1_d0100-tel_number,
smtp_addr_b TYPE sza1_d0100-smtp_addr,
street_b TYPE addr1_data-street,
city1_b TYPE addr1_data-city1,
region_b TYPE addr1_data-region,
post_code1_b TYPE addr1_data-post_code1,
country_b TYPE addr1_data-country,
matnr TYPE vbap-matnr,
maktx TYPE makt-maktx,
zmeng TYPE vbap-zmeng,
vbegdat TYPE veda-vbegdat,
venddat TYPE veda-venddat,
kzwi3 TYPE vbap-kzwi3,
con_val TYPE vbap-kzwi3,
sales_rep TYPE vbak-kunnr,
vkbur TYPE vbak-vkbur,
END OF x_final,
BEGIN OF x_log,
message TYPE string,
END OF x_log.
*/..Workarea Declarations
DATA: w_vbak TYPE ty_vbak,
w_veda TYPE ty_veda,
w_vbkd TYPE ty_vbkd,
w_vbap TYPE ty_vbap,
w_vbpa TYPE ty_vbpa,
w_adrc TYPE ty_adrc,
w_adr6 TYPE ty_adr6,
w_vbfa TYPE ty_vbfa,
w_vbeln TYPE ty_vbeln,
w_tvarv TYPE tvarv,
w_final TYPE x_final,
w_log TYPE x_log,
w_veda_new TYPE ty_veda,
w_kna1 TYPE ty_kna1.
*/..Data Declarations
DATA : gv_datefrm TYPE sy-datum,
gv_dateto TYPE sy-datum,
gv_order TYPE string.
DATA: t_vbak TYPE STANDARD TABLE OF ty_vbak,
t_vbkd TYPE STANDARD TABLE OF ty_vbkd,
t_veda TYPE STANDARD TABLE OF ty_veda,
t_vbap TYPE STANDARD TABLE OF ty_vbap,
t_vbpa TYPE STANDARD TABLE OF ty_vbpa,
t_vbpa_s TYPE STANDARD TABLE OF ty_vbpa_s,
t_adrc TYPE STANDARD TABLE OF ty_adrc,
t_adr6 TYPE STANDARD TABLE OF ty_adr6,
t_hrp1001 TYPE STANDARD TABLE OF ty_hrp1001,
t_but000 TYPE STANDARD TABLE OF ty_but000,
t_vbfa TYPE STANDARD TABLE OF ty_vbfa,
t_vbeln TYPE STANDARD TABLE OF ty_vbeln,
t_tvarv TYPE STANDARD TABLE OF tvarv,
t_final TYPE STANDARD TABLE OF x_final,
t_log TYPE STANDARD TABLE OF x_log,
t_veda_new TYPE STANDARD TABLE OF ty_veda,
t_kna1 TYPE STANDARD TABLE OF ty_kna1.
DATA: t_cdhdr TYPE TABLE OF cdhdr,
t_cdhdr1 TYPE TABLE OF cdhdr,
t_cdpos TYPE TABLE OF cdpos.
DATA: w_cdpos TYPE cdpos,
w_cdhdr TYPE cdhdr,
w_but000 TYPE ty_but000,
w_hrp1001 TYPE ty_hrp1001,
w_vbpa_s TYPE ty_vbpa_s.
DATA: r_beg_dat TYPE RANGE OF vbak-erdat,
r_canc_dat TYPE RANGE OF veda-vbedkue,
r_parvw TYPE RANGE OF vbpa-parvw,
r_auart TYPE RANGE OF vbak-auart,
rg_auart TYPE RANGE OF vbak-auart,
r_reinst TYPE RANGE OF vbak-erdat,
r_expired TYPE RANGE OF vbak-erdat,
r_contract TYPE RANGE OF vbak-erdat,
r_vkuegru TYPE RANGE OF veda-vkuegru,
r_mtart TYPE RANGE OF mara-mtart,
r_matnr TYPE RANGE OF mara-matnr.
DATA: w_beg_dat LIKE LINE OF r_beg_dat,
w_canc_dat LIKE LINE OF r_canc_dat,
w_parvw LIKE LINE OF r_parvw,
w_auart LIKE LINE OF r_auart,
wg_auart LIKE LINE OF rg_auart,
w_reinst LIKE LINE OF r_reinst,
w_expired LIKE LINE OF r_expired,
w_contract LIKE LINE OF r_contract,
w_vkuegru LIKE LINE OF r_vkuegru,
w_mtart LIKE LINE OF r_mtart.
*/..Selection screen declaration
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE text-001.
PARAMETERS: p_rname TYPE makt-maktx OBLIGATORY,
p_mid TYPE adr6-smtp_addr OBLIGATORY.
SELECT-OPTIONS: s_mat_i FOR mara-matnr,
s_mat_p FOR mara-matnr.
SELECTION-SCREEN END OF BLOCK b1.
* Frequency and Timing
SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE text-003.
PARAMETERS: p_daily RADIOBUTTON GROUP g1 DEFAULT 'X',
p_weekly RADIOBUTTON GROUP g1,
p_month RADIOBUTTON GROUP g1.
SELECTION-SCREEN END OF BLOCK b3.
*-Initialization-*
INITIALIZATION.
  PERFORM initialization.
*-Start-of-Selection-*
START-OF-SELECTION.
*/..Get the time period for which the report is run.
  PERFORM get_daterange.
*/..Get the Material types based on the selection.
  PERFORM sub_fill_material_types.
  PERFORM sub_consolidate_data.
*&---------------------------------------------------------------------*
*& Form INITIALIZATION
*&---------------------------------------------------------------------*
FORM initialization .
  wg_auart-sign = 'I'.
  wg_auart-option = 'EQ'.
  wg_auart-low = 'ZC01'.
  APPEND wg_auart TO rg_auart.
  CLEAR wg_auart.
  wg_auart-sign = 'I'.
  wg_auart-option = 'EQ'.
  wg_auart-low = 'ZC02'.
  APPEND wg_auart TO rg_auart.
  CLEAR wg_auart.
  wg_auart-sign = 'I'.
  wg_auart-option = 'EQ'.
  wg_auart-low = 'ZC03'.
  APPEND wg_auart TO rg_auart.
  CLEAR wg_auart.
  wg_auart-sign = 'I'.
  wg_auart-option = 'EQ'.
  wg_auart-low = 'ZR01'.
  APPEND wg_auart TO rg_auart.
  CLEAR wg_auart.
  wg_auart-sign = 'I'.
  wg_auart-option = 'EQ'.
  wg_auart-low = 'ZR02'.
  APPEND wg_auart TO rg_auart.
  CLEAR wg_auart.
ENDFORM. " INITIALIZATION
*&---------------------------------------------------------------------*
*& Form GET_DATERANGE
*&---------------------------------------------------------------------*
FORM get_daterange .
  DATA : lv_date TYPE sy-datum.
  CLEAR: lv_date,
  gv_datefrm,
  gv_dateto.
  IF p_daily EQ 'X'.
*/..Set current date
    gv_datefrm = sy-datum.
    gv_dateto = sy-datum.
  ELSEIF p_weekly EQ 'X'.
*/..get a date in previous week.
    CALL FUNCTION 'HR_99S_DATE_MINUS_TIME_UNIT'
      EXPORTING
        i_idate               = sy-datum
        i_time                = 7
        i_timeunit            = 'D'
      IMPORTING
        o_idate               = lv_date
      EXCEPTIONS
        invalid_period        = 1
        invalid_round_up_rule = 2
        internal_error        = 3
        OTHERS                = 4.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
*/..get the previous week's start and end date.
*/..by default monday would be the starting day and Sunday would be
* the ending day.
    CALL FUNCTION 'GET_WEEK_INFO_BASED_ON_DATE'
    EXPORTING
    date = lv_date
    IMPORTING
* WEEK =
    monday = gv_datefrm
    sunday = gv_dateto.
  ELSEIF p_month EQ 'X'.
*/..get a date in previous month.
    CALL FUNCTION 'HR_99S_DATE_MINUS_TIME_UNIT'
      EXPORTING
        i_idate               = sy-datum
        i_time                = 1
        i_timeunit            = 'M'
      IMPORTING
        o_idate               = lv_date
      EXCEPTIONS
        invalid_period        = 1
        invalid_round_up_rule = 2
        internal_error        = 3
        OTHERS                = 4.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
* /..Get the previous month's start and End Date.
    CALL FUNCTION 'OIL_MONTH_GET_FIRST_LAST'
      EXPORTING
        i_date      = lv_date
      IMPORTING
        e_first_day = gv_datefrm
        e_last_day  = gv_dateto
      EXCEPTIONS
        wrong_date  = 1
        OTHERS      = 2.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDIF.
ENDFORM. " GET_DATERANGE
*&---------------------------------------------------------------------*
*& Form GET_CONTRACT_LIST
*&---------------------------------------------------------------------*
FORM get_contract_list .
*/.. Get the Change history value for the given selection.
*/.. Change history might be Dialy, Weekly or Monthly.
*/..Based onthe checkbox selected pick the matching contract nos.
*/..for new orders.
  REFRESH r_beg_dat.
  w_beg_dat-sign = 'I'.
  w_beg_dat-option = 'BT'.
  w_beg_dat-low = gv_datefrm.
  w_beg_dat-high = gv_dateto.
  APPEND w_beg_dat TO r_beg_dat.
  CLEAR w_beg_dat.
*/..Get the contract header details.
  SELECT vbeln erdat auart augru vkbur kunnr
  FROM vbak
  INTO TABLE t_vbak
  WHERE erdat IN r_beg_dat
  AND auart IN r_auart.
  IF t_vbak[] IS NOT INITIAL.
*/..Get the Sales Document: Business Data details.
    SELECT vbeln posnr zzreinstate zzexpired zzlatesale zzrenewal
    FROM vbkd
    INTO TABLE t_vbkd
    FOR ALL ENTRIES IN t_vbak
    WHERE vbeln = t_vbak-vbeln.
*/.. Get the Contract data details
    SELECT vbeln vposn vbegdat venddat vkuegru vbedkue
    FROM veda
    INTO TABLE t_veda
    FOR ALL ENTRIES IN t_vbak
    WHERE vbeln = t_vbak-vbeln.
* AND vbedkue IN r_canc_dat
* AND vkuegru = gv_canc_reason.
*/..get material and its desc details from VBAP and MAKT table.
    PERFORM sub_get_item_details USING t_vbak.
*/..Get partner detials from VBPA table
* for the above obtained contracts.
*build a range holding only Ship-to,
* sold-to and Bill-to partner functions
    PERFORM sub_get_partner_details USING t_vbak.
*/.. Get the Address details for Sold to & Bill to partners
    PERFORM sub_get_partner_address_det USING t_vbpa.
*/.. Get the E-mail addresses for Sold to & Bill to addresses
    PERFORM sub_get_email_address USING t_vbpa.
*/.. Get Sales Rep Details for all the Contract documents
    PERFORM sub_get_sales_rep_details USING t_vbpa_s.
  ENDIF.
ENDFORM. " GET_CONTRACT_LIST
*&---------------------------------------------------------------------*
*& Form SUB_CONSOLIDATE_DATA_NEW
*&---------------------------------------------------------------------*
FORM sub_consolidate_data_new.
  LOOP AT t_vbak INTO w_vbak WHERE auart IN r_auart.
    w_final-vbeln = w_vbak-vbeln.
    w_final-auart = w_vbak-auart.
    w_final-erdat = w_vbak-erdat.
    w_final-augru = w_vbak-augru.
    w_final-vkbur = w_vbak-vkbur.
*/.. Sold to Party addresses
    CLEAR w_vbpa.
    READ TABLE t_vbpa INTO w_vbpa WITH KEY vbeln = w_vbak-vbeln parvw = 'AG'.
    IF sy-subrc = 0.
      w_final-kunnr = w_vbpa-kunnr.
      READ TABLE t_adrc INTO w_adrc WITH KEY addrnumber = w_vbpa-adrnr.
      IF sy-subrc = 0.
        CONCATENATE w_adrc-name1 w_adrc-name2 INTO w_final-name SEPARATED BY space.
        w_final-telf1 = w_adrc-tel_number.
        w_final-street = w_adrc-street.
        w_final-city1 = w_adrc-city1.
        w_final-region = w_adrc-region.
        w_final-post_code1 = w_adrc-post_code1.
        w_final-country = w_adrc-country.
      ENDIF.
      READ TABLE t_adr6 INTO w_adr6 WITH KEY addrnumber = w_vbpa-adrnr.
      IF sy-subrc = 0.
        w_final-smtp_addr = w_adr6-smtp_addr.
      ENDIF.
    ENDIF.
*/.. Bill to party address
    CLEAR w_vbpa.
    READ TABLE t_vbpa INTO w_vbpa WITH KEY vbeln = w_vbak-vbeln parvw = 'RE'.
    IF sy-subrc = 0.
      w_final-kunnr_b = w_vbpa-kunnr.
      CLEAR w_adrc.
      READ TABLE t_adrc INTO w_adrc WITH KEY addrnumber = w_vbpa-adrnr.
      IF sy-subrc = 0.
        CONCATENATE w_adrc-name1 w_adrc-name2 INTO w_final-name_b SEPARATED BY space.
        w_final-telf1_b = w_adrc-tel_number.
        w_final-street_b = w_adrc-street.
        w_final-city1_b = w_adrc-city1.
        w_final-region_b = w_adrc-region.
        w_final-post_code1_b = w_adrc-post_code1.
        w_final-country_b = w_adrc-country.
      ENDIF.
      CLEAR w_adr6.
      READ TABLE t_adr6 INTO w_adr6 WITH KEY addrnumber = w_vbpa-adrnr.
      IF sy-subrc = 0.
        w_final-smtp_addr_b = w_adr6-smtp_addr.
      ENDIF.
    ENDIF.
* Line Item Details
    LOOP AT t_vbap INTO w_vbap WHERE vbeln = w_vbak-vbeln.
      w_final-posnr = w_vbap-posnr.
      w_final-matnr = w_vbap-matnr.
      w_final-maktx = w_vbap-maktx.
      w_final-zmeng = w_vbap-zmeng.
      w_final-kzwi3 = w_vbap-kzwi3 + w_vbap-kzwi5.
      READ TABLE t_vbpa INTO w_vbpa WITH KEY vbeln = w_vbak-vbeln posnr = w_vbap-posnr parvw = 'Z3'.
      IF sy-subrc = 0.
        CLEAR w_hrp1001.
        READ TABLE t_hrp1001 INTO w_hrp1001 WITH KEY sobid = w_vbpa-kunnr. "#EC *
        IF sy-subrc = 0.
          w_final-sales_rep = w_hrp1001-sobid.
        ENDIF.
      ELSE.
        CLEAR w_vbpa.
        READ TABLE t_vbpa INTO w_vbpa WITH KEY vbeln = w_vbak-vbeln parvw = 'Z3'.
        IF sy-subrc = 0.
          CLEAR w_hrp1001.
          READ TABLE t_hrp1001 INTO w_hrp1001 WITH KEY sobid = w_vbpa-kunnr. "#EC *
          IF sy-subrc = 0.
            w_final-sales_rep = w_hrp1001-sobid.
          ENDIF.
        ENDIF.
      ENDIF.
      CLEAR w_veda.
      READ TABLE t_veda INTO w_veda WITH KEY vbeln = w_vbap-vbeln vposn = w_vbap-posnr.
      IF sy-subrc = 0.
        w_final-vbegdat = w_veda-vbegdat.
        w_final-venddat = w_veda-venddat.
      ELSE.
        READ TABLE t_veda INTO w_veda WITH KEY vbeln = w_vbap-vbeln." vposn = w_vbap-posnr.
        IF sy-subrc = 0.
          w_final-vbegdat = w_veda-vbegdat.
          w_final-venddat = w_veda-venddat.
        ENDIF.
      ENDIF.
      APPEND w_final TO t_final.
    ENDLOOP.
    CLEAR w_final.
    CLEAR w_vbap.
    CLEAR w_vbak.
  ENDLOOP.
  REFRESH: t_vbak,
  t_veda,
  t_vbkd,
  t_vbap,
  t_vbpa,
  t_adr6,
  t_adrc.
ENDFORM. " SUB_CONSOLIDATE_DATA_NEW
*&---------------------------------------------------------------------*
*& Form SUB_SEND_MAIL
*&---------------------------------------------------------------------*
FORM sub_send_mail.
* Data Declaration for Mail Sending Options
* Receiver
  DATA: l_recipent TYPE REF TO if_recipient_bcs,
* Sender
  l_sender TYPE REF TO cl_sapuser_bcs,
  l_attcdoctype TYPE soodk-objtp,
  l_atttitle TYPE sood-objdes,
  l_freq TYPE string,
  l_from TYPE char10,
  l_to TYPE char10.
* Execptions
  DATA : l_bcs_exception TYPE REF TO cx_bcs,               "#EC NEEDED)
  l_document TYPE REF TO cl_document_bcs,
  l_send_request TYPE REF TO cl_bcs.
  DATA: t_mail_text TYPE bcsy_text,
  wa_mail_text_row TYPE soli,
  l_subject2 TYPE so_obj_des,
  l_message TYPE string,
  l_message1 TYPE string.
  DATA: l_num_rows TYPE i,
  l_text_length TYPE so_obj_len,
  l_num_line TYPE i.
  DATA: l_line TYPE string,
  l_zmeng TYPE string,
  l_kzwi3 TYPE string.
  DATA: l_venddat TYPE char10,
  l_vbegdat TYPE char10,
  l_erdat TYPE char10.
  DATA: gt_binary_content TYPE solix_tab,
  gv_size TYPE so_obj_len.
  CONSTANTS: c_tab TYPE c VALUE cl_bcs_convert=>gc_tab,
  c_cr TYPE c VALUE cl_bcs_convert=>gc_crlf,
  c_ext TYPE soodk-objtp VALUE 'XLS',
  c_x TYPE c VALUE 'X'.
  TYPES : BEGIN OF ty_tvakt,
  auart TYPE auart,
  bezei TYPE bezei20,
  END OF ty_tvakt,
  BEGIN OF ty_tvaut,
  augru TYPE augru,
  bezei TYPE bezei40,
  END OF ty_tvaut.
  DATA : lt_tvakt TYPE STANDARD TABLE OF ty_tvakt,
  wa_tvakt TYPE ty_tvakt,
  lt_tvaut TYPE STANDARD TABLE OF ty_tvaut,
  wa_tvaut TYPE ty_tvaut,
  lt_ordesc TYPE STANDARD TABLE OF x_final,
  lv_auart TYPE auart,
  lv_augru TYPE augru,
  lv_auarttxt TYPE bezei40,
  lv_augrutxt TYPE bezei40,
  c_hyp(3) TYPE c VALUE ' - '.
  IF NOT t_final IS INITIAL.
    lt_ordesc = t_final.
    SORT lt_ordesc BY auart augru.
    DELETE ADJACENT DUPLICATES FROM lt_ordesc COMPARING auart augru.
    CLEAR : lt_tvakt,
    lt_tvaut,
    lv_auart,
    lv_augru.
    SELECT auart bezei INTO TABLE lt_tvakt
    FROM tvakt
    FOR ALL ENTRIES IN lt_ordesc
    WHERE auart EQ lt_ordesc-auart
    AND spras EQ sy-langu.
    SORT lt_tvakt BY auart.
    DELETE ADJACENT DUPLICATES FROM lt_tvakt COMPARING ALL FIELDS.
    SELECT augru bezei INTO TABLE lt_tvaut
    FROM tvaut
    FOR ALL ENTRIES IN lt_ordesc
    WHERE augru EQ lt_ordesc-augru
    AND spras EQ sy-langu.
    SORT lt_tvaut BY augru.
    DELETE ADJACENT DUPLICATES FROM lt_tvaut COMPARING ALL FIELDS.
  ENDIF.
* Subject - Header
  l_subject2 = p_rname.
  CONCATENATE gv_datefrm+4(2) '/' gv_datefrm+6(2) '/' gv_datefrm+0(4) INTO l_from.
  CONCATENATE gv_dateto+4(2) '/' gv_dateto+6(2) '/' gv_dateto+0(4) INTO l_to.
  CLEAR l_freq.
  IF p_daily = c_x.
    l_freq = 'Daily'.
  ELSEIF p_weekly = c_x.
    l_freq = 'Weekly'.
  ELSEIF p_month = c_x.
    l_freq = 'Monthly'.
  ENDIF.
  DESCRIBE TABLE t_final LINES l_num_line.
  TRY.
      wa_mail_text_row = 'Data Selection Details:'.
      APPEND wa_mail_text_row TO t_mail_text. CLEAR wa_mail_text_row.
      APPEND wa_mail_text_row TO t_mail_text. CLEAR wa_mail_text_row.
      CONCATENATE 'Contract Type' ':' gv_order INTO wa_mail_text_row SEPARATED BY space.
      APPEND wa_mail_text_row TO t_mail_text.
      CLEAR wa_mail_text_row.
      CONCATENATE 'Frequency' ':' l_freq INTO wa_mail_text_row SEPARATED BY space.
      APPEND wa_mail_text_row TO t_mail_text.
      CLEAR wa_mail_text_row.
      CONCATENATE 'Date' ':' l_from 'to' l_to INTO wa_mail_text_row SEPARATED BY space.
      APPEND wa_mail_text_row TO t_mail_text. CLEAR wa_mail_text_row.
      CLEAR l_message.
      l_message = l_num_line.
      CONCATENATE 'No. of Records' ':' l_message INTO wa_mail_text_row SEPARATED BY space.
      APPEND wa_mail_text_row TO t_mail_text. CLEAR wa_mail_text_row.
      APPEND wa_mail_text_row TO t_mail_text.
      APPEND wa_mail_text_row TO t_mail_text.
      IF t_final[] IS INITIAL.
        wa_mail_text_row = 'No document exists for this selection.'.
        APPEND wa_mail_text_row TO t_mail_text.
        CLEAR wa_mail_text_row.
      ENDIF.
* Define rows and file size
      DESCRIBE TABLE t_mail_text LINES l_num_rows.
      l_num_rows = l_num_rows * 255.
      MOVE l_num_rows TO l_text_length.
      TRY.
          CALL METHOD cl_document_bcs=>create_document
            EXPORTING
              i_type    = 'RAW'
              i_subject = l_subject2
              i_length  = l_text_length
              i_text    = t_mail_text
            RECEIVING
              result    = l_document.
        CATCH cx_document_bcs .                         "#EC NO_HANDLER
      ENDTRY.
* IF t_final[] IS NOT INITIAL.
* Create Attachment
      CONCATENATE c_tab
      text-011 c_tab
      text-012 c_tab
      text-013 c_tab
      text-014 c_tab
      text-015 c_tab
      text-016 c_tab
      text-017 c_tab
      text-018 c_tab
      text-019 c_tab
      text-020 c_tab
      text-021 c_tab
      text-022 c_tab
      text-023 c_tab
      text-024 c_tab
      text-025 c_tab
      text-026 c_tab
      text-027 c_tab
      text-028 c_tab
      text-029 c_tab
      text-030 c_tab
      text-031 c_tab
      text-032 c_tab
      text-033 c_tab
      text-034 c_tab
      text-035 c_tab
      text-036 c_tab
      text-037 c_tab
      text-038 c_tab
      text-039 c_cr INTO l_line.
      CLEAR w_final.
      LOOP AT t_final INTO w_final.
        l_zmeng = w_final-zmeng.
        l_kzwi3 = w_final-kzwi3.
        CLEAR: l_vbegdat,
        l_venddat,
        l_erdat.
        CONCATENATE w_final-vbegdat+4(2) '/'
        w_final-vbegdat+6(2) '/'
        w_final-vbegdat+0(4)
        INTO l_vbegdat.
        CONCATENATE w_final-venddat+4(2) '/'
        w_final-venddat+6(2) '/'
        w_final-venddat+0(4)
        INTO l_venddat.
        CONCATENATE w_final-erdat+4(2) '/'
        w_final-erdat+6(2) '/'
        w_final-erdat+0(4)
        INTO l_erdat.
        IF w_final-auart NE lv_auart.
          READ TABLE lt_tvakt INTO wa_tvakt
          WITH KEY auart = w_final-auart.
          IF sy-subrc EQ 0.
            lv_auarttxt = wa_tvakt-bezei.
            CONCATENATE w_final-auart lv_auarttxt INTO lv_auarttxt SEPARATED BY c_hyp.
          ELSE.
            lv_auarttxt = w_final-auart.
          ENDIF.
          lv_auart = w_final-auart.
        ENDIF.
        IF w_final-augru NE lv_augru.
          READ TABLE lt_tvaut INTO wa_tvaut
          WITH KEY augru = w_final-augru.
          IF sy-subrc EQ 0.
            lv_augrutxt = wa_tvaut-bezei.
            CONCATENATE w_final-augru lv_augrutxt INTO lv_augrutxt SEPARATED BY c_hyp.
          ELSE.
            lv_augrutxt = w_final-augru.
          ENDIF.
          lv_augru = w_final-augru.
        ENDIF.
        CONCATENATE l_line
        lv_auarttxt
        l_erdat
        lv_augrutxt
        w_final-vbeln
        w_final-posnr
        w_final-kunnr
        w_final-name
        w_final-telf1
        w_final-smtp_addr
        w_final-street
        w_final-city1
        w_final-country
        w_final-post_code1
        w_final-region
        w_final-kunnr_b
        w_final-name_b
        w_final-street_b
        w_final-city1_b
        w_final-region_b
        w_final-post_code1_b
        w_final-telf1_b
        w_final-matnr
        w_final-maktx
        l_zmeng
        l_vbegdat
        l_venddat
        l_kzwi3
        w_final-sales_rep
        w_final-vkbur
        INTO l_line SEPARATED BY c_tab.
        CONCATENATE l_line c_cr INTO l_line.
        CLEAR w_final.
      ENDLOOP.
      TRY.
          cl_bcs_convert=>string_to_solix(
          EXPORTING
          iv_string = l_line
          iv_codepage = '4103' "suitable for MS Excel, leave empty
          iv_add_bom = 'X' "for other doc types
          IMPORTING
          et_solix = gt_binary_content
          ev_size = gv_size ).
        CATCH cx_bcs.
          MESSAGE e445(so).
      ENDTRY.
* ENDIF.
* Define File Name
      l_attcdoctype = c_ext.
      l_atttitle = p_rname.
* Create Document
      CALL METHOD l_document->add_attachment(
        i_attachment_type = l_attcdoctype
        i_attachment_subject = l_atttitle
        i_attachment_size = gv_size
        i_att_content_hex = gt_binary_content ).
      l_send_request = cl_bcs=>create_persistent( ).
      l_send_request->set_document( l_document ).
* Define Sender
      l_sender = cl_sapuser_bcs=>create( sy-uname ).
      TRY.
          CALL METHOD l_send_request->set_sender
            EXPORTING
              i_sender = l_sender.
        CATCH cx_send_req_bcs .                         "#EC NO_HANDLER
      ENDTRY.
* Define Recipient
      l_recipent = cl_cam_address_bcs=>create_internet_address( p_mid ).
      l_send_request->add_recipient( EXPORTING i_recipient = l_recipent ).
* Schedule
      l_send_request->set_send_immediately( 'X' ).
      l_send_request->send( ).
      COMMIT WORK AND WAIT.
* Catch Execptions
    CATCH cx_bcs INTO l_bcs_exception.                  "#EC NO_HANDLER
  ENDTRY.
  IF sy-subrc = 0.
    CONCATENATE gv_order '-' l_message INTO l_message1 SEPARATED BY space.
    w_log-message = l_message1.
    APPEND w_log TO t_log.
    CLEAR w_log.
  ENDIF.
  REFRESH t_final.
  CLEAR w_final.
ENDFORM. " SUB_SEND_MAIL
*&---------------------------------------------------------------------*
*& Form SUB_CONSOLIDATE_DATA
*&---------------------------------------------------------------------*
FORM sub_consolidate_data.
*/.. Consolidate and send the New Orders
  CLEAR gv_order.
  gv_order = 'New Contracts'.
*/..Get the list of document types to fetch the contracts.
  PERFORM get_document_types_new.
*/..Based on the selected check boxes, pick the required details.
  PERFORM get_contract_list.
*/.. Consolidate New orders
  PERFORM sub_consolidate_data_new.
*/.. Send New orders
  PERFORM sub_send_mail.
* Display final Log
  PERFORM sub_display_log.
ENDFORM. " SUB_CONSOLIDATE_DATA
*&---------------------------------------------------------------------*
*& Form SUB_GET_ITEM_DETAILS
*&---------------------------------------------------------------------*
FORM sub_get_item_details USING t_vbak LIKE t_vbak.
  REFRESH t_vbap.
*/..get material and its desc details from VBAP and MAKT table.
  IF NOT t_vbak[] IS INITIAL.
    SELECT a~vbeln a~posnr a~matnr a~zmeng a~kzwi3 a~kzwi5
    b~maktx c~mtart
    FROM vbap AS a INNER JOIN
    makt AS b ON a~matnr = b~matnr
    INNER JOIN mara AS c ON c~matnr = b~matnr
    INTO TABLE t_vbap
    FOR ALL ENTRIES IN t_vbak
    WHERE a~vbeln = t_vbak-vbeln
    AND c~matnr IN r_matnr
    AND c~mtart IN r_mtart.
  ENDIF.
ENDFORM. " SUB_GET_ITEM_DETAILS
*&---------------------------------------------------------------------*
*& Form SUB_GET_PARTNER_DETAILS
*&---------------------------------------------------------------------*
FORM sub_get_partner_details USING t_vbak LIKE t_vbak.
  REFRESH: r_parvw,
  t_vbpa,
  t_vbpa_s.
  w_parvw-sign = 'I'.
  w_parvw-option = 'EQ'.
  w_parvw-low = 'AG'.
  APPEND w_parvw TO r_parvw.
  CLEAR w_parvw.
  w_parvw-sign = 'I'.
  w_parvw-option = 'EQ'.
  w_parvw-low = 'RE'.
  APPEND w_parvw TO r_parvw.
  CLEAR w_parvw.
  w_parvw-sign = 'I'.
  w_parvw-option = 'EQ'.
  w_parvw-low = 'SH'.
  APPEND w_parvw TO r_parvw.
  CLEAR w_parvw.
  w_parvw-sign = 'I'.
  w_parvw-option = 'EQ'.
  w_parvw-low = 'Z3'.
  APPEND w_parvw TO r_parvw.
  CLEAR w_parvw.
  SELECT vbeln posnr parvw kunnr pernr adrnr
  FROM vbpa INTO TABLE t_vbpa
  FOR ALL ENTRIES IN t_vbak
  WHERE vbeln = t_vbak-vbeln
  AND parvw IN r_parvw.
  LOOP AT t_vbpa INTO w_vbpa WHERE parvw = 'Z3'.
    w_vbpa_s-sobid = w_vbpa-kunnr.
    w_vbpa_s-kunnr = w_vbpa-kunnr.
    APPEND w_vbpa_s TO t_vbpa_s.
    CLEAR w_vbpa_s.
  ENDLOOP.
ENDFORM. " SUB_GET_PARTNER_DETAILS
*&---------------------------------------------------------------------*
*& Form SUB_GET_PARTNER_ADDRESS_DET
*&---------------------------------------------------------------------*
FORM sub_get_partner_address_det USING t_vbpa LIKE t_vbpa.
  REFRESH t_adrc.
  IF t_vbpa[] IS NOT INITIAL.
    SELECT addrnumber name1 name2 street city1 post_code1 region country tel_number
    FROM adrc
    INTO TABLE t_adrc
    FOR ALL ENTRIES IN t_vbpa
    WHERE addrnumber = t_vbpa-adrnr.
  ENDIF.
ENDFORM. " SUB_GET_PARTNER_ADDRESS_DET
*&---------------------------------------------------------------------*
*& Form SUB_GET_EMAIL_ADDRESS
*&---------------------------------------------------------------------*
FORM sub_get_email_address USING t_vbpa LIKE t_vbpa.
  REFRESH t_adr6.
  IF t_vbpa[] IS NOT INITIAL.
    SELECT addrnumber smtp_addr
    FROM adr6
    INTO TABLE t_adr6
    FOR ALL ENTRIES IN t_vbpa
    WHERE addrnumber = t_vbpa-adrnr.
  ENDIF.
ENDFORM. " SUB_GET_EMAIL_ADDRESS
*&---------------------------------------------------------------------*
*& Form SUB_GET_SALES_REP_DETAILS
*&---------------------------------------------------------------------*
FORM sub_get_sales_rep_details USING t_vbpa_s LIKE t_vbpa_s.
  REFRESH: t_but000,
  t_hrp1001.
  SELECT otype objid plvar relat begda endda sclas sobid
  FROM hrp1001
  INTO TABLE t_hrp1001
  FOR ALL ENTRIES IN t_vbpa_s
  WHERE sobid = t_vbpa_s-sobid
  AND otype = 'S'
  AND plvar = '01'
  AND relat = '008'
  AND sclas = 'BP'
  AND begda < sy-datum
  AND endda >= sy-datum.
  SELECT partner bpkind bu_group name_last name_first
  FROM but000 INTO TABLE t_but000
  FOR ALL ENTRIES IN t_hrp1001
  WHERE partner = t_hrp1001-sobid+0(10)
  AND bpkind = '9002'.
ENDFORM. " SUB_GET_SALES_REP_DETAILS
*&---------------------------------------------------------------------*
*& Form SUB_FILL_MATERIAL_TYPES
*&---------------------------------------------------------------------*
FORM sub_fill_material_types .
  REFRESH r_mtart.
  REFRESH r_matnr.
  IF s_mat_i IS NOT INITIAL.
    APPEND LINES OF s_mat_i TO r_matnr.
  ENDIF.
  IF s_mat_p IS NOT INITIAL.
    APPEND LINES OF s_mat_p TO r_matnr.
  ENDIF.
ENDFORM. " SUB_FILL_MATERIAL_TYPES
*&---------------------------------------------------------------------*
*& Form GET_DOCUMENT_TYPES_NEW
*&---------------------------------------------------------------------*
FORM get_document_types_new .
  REFRESH r_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC01'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC02'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC03'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
ENDFORM. " GET_DOCUMENT_TYPES_NEW
*&---------------------------------------------------------------------*
*& Form GET_DOCUMENT_TYPES_RENEW
*&---------------------------------------------------------------------*
FORM get_document_types_renew .
  REFRESH r_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZR01'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZR02'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
ENDFORM. " GET_DOCUMENT_TYPES_RENEW
*&---------------------------------------------------------------------*
*& Form SUB_DISPLAY_LOG
*&---------------------------------------------------------------------*
FORM sub_display_log .
  DATA: l_mesg TYPE string.
  IF t_log IS NOT INITIAL.
    CONCATENATE 'Email sent to' p_mid 'successfully' INTO l_mesg SEPARATED BY space.
  ELSE.
    CONCATENATE 'Email failed to sent -' p_mid INTO l_mesg SEPARATED BY space.
  ENDIF.
  WRITE / l_mesg.
  LOOP AT t_log INTO w_log.
    WRITE: / w_log-message.
    CLEAR w_log.
  ENDLOOP.
ENDFORM. " SUB_DISPLAY_LOG
*&---------------------------------------------------------------------*
*& Form SUB_VALIDATE_PARTNER_DATA
*&---------------------------------------------------------------------*
FORM sub_validate_partner_data .
  REFRESH t_cdhdr1[].
  SORT t_cdhdr BY objectid.
  t_cdhdr1[] = t_cdhdr[].
  LOOP AT t_cdhdr1 INTO w_cdhdr WHERE change_ind = 'I'.
    READ TABLE t_cdhdr1 TRANSPORTING NO FIELDS WITH KEY objectid = w_cdhdr-objectid
    udate = w_cdhdr-udate
    change_ind = 'U'.
    IF sy-subrc = 0.
      DELETE t_cdhdr WHERE objectid = w_cdhdr-objectid.
    ENDIF.
    CLEAR w_cdhdr.
  ENDLOOP.
ENDFORM. " SUB_VALIDATE_PARTNER_DATA
*&---------------------------------------------------------------------*
*& Form SUB_GET_DOCUMENT_TYPE_ADD
*&---------------------------------------------------------------------*
FORM sub_get_document_type_add .
  REFRESH r_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC01'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC02'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC03'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC04'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC05'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZC06'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZR01'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
  w_auart-sign = 'I'.
  w_auart-option = 'EQ'.
  w_auart-low = 'ZR02'.
  APPEND w_auart TO r_auart.
  CLEAR w_auart.
ENDFORM. " SUB_GET_DOCUMENT_TYPE_ADD
