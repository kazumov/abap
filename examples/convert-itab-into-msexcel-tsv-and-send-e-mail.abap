*&---------------------------------------------------------------------*
*& Report Z_EXCEL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report z_excel.

class cl_excel_tsv definition.
  public section.
    methods: itab_to_tsv
      importing itab  type any table
      changing
        t_tsv type string_table
        t_tsv_solix type solix_tab.

  private section.
    methods:
      to_string
        importing dat                 type data
        returning value(string_value) type string,
      get_type
        importing dat             type data
        returning value(dat_type) type string,
      d_to_excel_date
        importing dat               type d
        returning value(excel_date) type string,
      f_to_excel_number
        importing dat                 type f
        returning value(excel_number) type string,
      t_to_excel_time
        importing dat               type t
        returning value(excel_time) type string,
      p_to_excel_number
        importing dat                 type p
        returning value(excel_number) type string,
      i_to_excel_number
        importing dat                 type i
        returning value(excel_number) type string.
endclass.

class cl_excel_tsv implementation.
  method itab_to_tsv.
    " read the structure of itab
    data: r_w_itab type ref to data,
          r_des    type ref to cl_abap_structdescr,
          t_comp   type abap_component_tab,
          t_row    type table of string,
          row      type string.

    field-symbols: <cell> type any.
    create data r_w_itab like line of itab.
    r_des ?= cl_abap_typedescr=>describe_by_data_ref( r_w_itab ).
    t_comp = r_des->get_components( ).

    "@todo create header
      loop at t_comp assigning field-symbol(<header_component>).
        append <header_component>-name to t_row.
      endloop.
      concatenate lines of t_row into row separated by cl_abap_char_utilities=>horizontal_tab.
      concatenate row cl_abap_char_utilities=>cr_lf into row.
      append row to t_tsv.
    clear: t_row, row.
    refresh t_row.

    " create body
    loop at itab assigning field-symbol(<record>).
      loop at t_comp assigning field-symbol(<component>).
        assign component <component>-name of structure <record> to <cell>.
        append me->to_string( <cell> ) to t_row.
      endloop.
      concatenate lines of t_row into row separated by cl_abap_char_utilities=>horizontal_tab.
      concatenate row cl_abap_char_utilities=>cr_lf into row.
      append row to t_tsv.
      clear: row, t_row.
      refresh t_row.
    endloop.
    " encode
    data: str_tab type string,
          r_conv_exc type ref to cx_bcs .
    concatenate lines of t_tsv into str_tab.

    try.
    call method cl_bcs_convert=>string_to_solix
      exporting
        iv_string   = str_tab
        iv_codepage = '1160'
        iv_add_bom  = 'X'
      importing
        et_solix    = t_tsv_solix
*        ev_size     =
        .
     catch cx_bcs into r_conv_exc.
    endtry.

  endmethod.
  method to_string.
    data: dat_type    type string,
          r_exc       type ref to cx_root.

    dat_type = me->get_type( dat ).

    if dat_type = 'D'.
      string_value = me->d_to_excel_date( dat ).
    elseif dat_type = 'F'.
      string_value = me->f_to_excel_number( dat ).
    elseif dat_type = 'T'.
      string_value = me->t_to_excel_time( dat ).
    elseif dat_type = 'P'.
      string_value = me->p_to_excel_number( dat ).
    elseif dat_type = 'I'.
      string_value = me->i_to_excel_number( dat ).
    elseif dat_type = 'C' or dat_type = 'g'.
      string_value = dat.
      replace all occurrences of regex '\t' in string_value with ` `.
    else.
      string_value = 'Can not display the values of type ' && dat_type && '.'.
    endif.
  endmethod.
  method get_type.
    data: r_des type ref to cl_abap_datadescr.
    r_des ?= cl_abap_typedescr=>describe_by_data( dat ).
    dat_type = r_des->type_kind.
  endmethod.
  method d_to_excel_date.
    concatenate '"' dat(4) '-' dat+4(2) '-' dat+6(2) '"' into excel_date.
  endmethod.
  method f_to_excel_number.
    if dat < 0.
        excel_number = abs( dat ).
        concatenate '-' excel_number into excel_number.
      else.
        excel_number = dat.
      endif.
  endmethod.
  method t_to_excel_time.
    concatenate '"' dat(2) ':' dat+2(2) ':' dat+4(2) '"' into excel_time.
  endmethod.
  method p_to_excel_number.
    data: r_exc type ref to cx_root,
          float_value type f.
    try.
          float_value = dat.
          excel_number = me->f_to_excel_number( float_value ).
        catch cx_root into r_exc.
          excel_number = ' %( '.
      endtry.
  endmethod.
  method i_to_excel_number.
    if dat < 0.
        excel_number = abs( dat ).
        concatenate '-' excel_number into excel_number.
      else.
        excel_number = dat.
      endif.
  endmethod.
endclass.



start-of-selection.
* PART 1
* Data preparation
*
  types: orders type zorders. " set any sap table
  data: w_orders type orders,
        t_orders type table of orders.

  select * from zorders into table t_orders up to 5000 rows. " set the sap table

  data: r_tsv type ref to cl_excel_tsv,
        t_tsv type table of string,
         t_solix type solix_tab.

  create object r_tsv.
  r_tsv->itab_to_tsv(
    exporting itab = t_orders
    changing t_tsv = t_tsv
      t_tsv_solix = t_solix ).


clear t_tsv.
refresh t_tsv.
clear t_solix.
refresh t_solix.



  types: begin of demo,
    i0 type i,
    d1 type d,
    p2 type p,
    f3 type f,
    t4 type t,
    c5 type string,
    end of demo.
   data: t_demo type table of demo,
         w_demo type demo.

   w_demo = value demo( i0 = '4-' d1 = '20150422' p2 = '3904.003-' f3 = '436.34'  t4 = '221355' c5 = 'tab_tab' ).
  append w_demo to t_demo.
   w_demo = value demo( i0 = '3' d1 = '20160321' p2 = '3904003.8-' f3 = '4.3634'  t4 = '121605' c5 = 'tab' && cl_abap_char_utilities=>horizontal_tab && 'tab' ).
  append w_demo to t_demo.


  r_tsv->itab_to_tsv(
    exporting itab = t_demo
    changing t_tsv = t_tsv t_tsv_solix = t_solix ).

* PART 5
* Send the message with the file attachment
*
  class cl_bcs definition load.
  data:  lo_send_request type ref to cl_bcs
        ,lo_document     type ref to cl_document_bcs
        ,lo_sender       type ref to if_sender_bcs
        ,lo_recipient    type ref to if_recipient_bcs
        ,lt_message_body type bcsy_text
        ,lx_document_bcs type ref to cx_document_bcs
        ,lv_send         type ad_smtpadr value 'ruben@kazumov.com'
        ,lv_sent_to_all  type os_boolean
        .
  "create send request
  lo_send_request = cl_bcs=>create_persistent( ).

  "create message body and subject
  append 'Dear Friend,' to lt_message_body.
  append initial line to lt_message_body.
  append 'Please review the attached Microsoft Excel TSV file.' to lt_message_body.
  append initial line to lt_message_body.
  append 'Thank You,' to lt_message_body.

  "put your text into the document
  lo_document = cl_document_bcs=>create_document(
                   i_type = 'RAW'
                   i_text = lt_message_body
                   i_subject = 'Internal table' ).


  try.
      lo_document->add_attachment(
        exporting
          i_attachment_type = 'TSV'
          i_attachment_subject = 'table_test_convert'
          i_att_content_hex = t_solix "encoded_content
          ).

    catch cx_document_bcs into lx_document_bcs.
      cl_demo_output=>display_text( 'error of send' ).
      " can not create the document
  endtry.

* Add attachment
* Pass the document to send request
  lo_send_request->set_document( lo_document ).


  "Create sender
  lo_sender = cl_cam_address_bcs=>create_internet_address( lv_send ).

  "Set sender
  lo_send_request->set_sender( lo_sender ).

  "Create recipient
  lo_recipient = cl_sapuser_bcs=>create( sy-uname ).
  "lo_recipient = cl_cam_address_bcs=>create_internet_address('ruben@kazumov.com' ). " set the e-mail address

*Set recipient
  lo_send_request->add_recipient(
       exporting
         i_recipient = lo_recipient i_express = 'X' ).

  lo_send_request->add_recipient( lo_recipient ).

* Set time send
*lo_send_request->set_send_immediately( 'X' ).


* Send email
  lo_send_request->send(
    exporting
      i_with_error_screen = 'X'
    receiving
      result = lv_sent_to_all ).

  commit work.
