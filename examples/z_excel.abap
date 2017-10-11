*&---------------------------------------------------------------------*
*& Report Z_EXCEL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report z_excel.

* THE SCRIPT CUSTOMIZING
* In line 19 - set/change the name of sap table. Current value is "ZORDERS"
* In line 23 - set/change the name of the sap table. Current value is "ZORDERS"
* In line 89 - set/change the filename. Current value is "C:\Users\r\Desktop\table_test_convert.xls"
* In line 158 - set/change the e-mail of sender.
* In lines 202 and 203 - choose the type of the report delivery local/e-mail
* In line 203 - set/change e-mail address

* PART 1
* Data preparation
*
types: orders type zorders. " set any sap table
data: w_orders type orders,
      t_orders type table of orders.

select * from zorders into table t_orders up to 10 rows. " set the sap table

* PART 2
* Convert original structure to one column string table
* with the data fields, separated by TAB and CRLF at the
* end of lines.
*
data: r_des             type ref to cl_abap_structdescr,
      r_orders          type ref to data,
      t_comp            type abap_component_tab,
      r_exc             type ref to cx_root,
      t_row             type table of string,
      tab_separated_row type string, " @old: type soli
      t_lines           type table of string, " @old: type table of soli
      table_as_string   type string,
      cast_string       type string.

field-symbols: <cell> type any.

create data r_orders like line of t_orders.
r_des ?= cl_abap_typedescr=>describe_by_data_ref( r_orders ).
t_comp = r_des->get_components( ).

loop at t_orders assigning field-symbol(<order>).
  loop at t_comp assigning field-symbol(<component>).
    assign component <component>-name of structure <order> to <cell>.
    cast_string = <cell>.
    append cast_string to t_row.
  endloop.
  append cl_abap_char_utilities=>cr_lf to t_row.
  concatenate lines of t_row into tab_separated_row separated by cl_abap_char_utilities=>horizontal_tab.
  append tab_separated_row to t_lines.
  clear t_row.
endloop.
concatenate lines of t_lines into table_as_string.
unassign: <cell>, <component>, <order>.
clear: tab_separated_row, t_orders, t_comp, t_lines.

* PART 3
* Create the MS EXCEL file body binary content
*
data: encoded_content        type solix_tab,
      encoded_content_length type so_obj_len,
      convertion_error       type ref to cx_bcs.
try.
    call method cl_bcs_convert=>string_to_solix
      exporting
        iv_string   = table_as_string
        iv_codepage = '1160' " 1160 = windows-1252; 4103 = utf-16le; 4110 = utf-8; 1105 = US-ASCII (7 bits); 1100 = iso-8859-1; 4102 = utf-16be
        iv_add_bom  = 'X'
      importing
        et_solix    = encoded_content
        ev_size     = encoded_content_length.
  catch cx_bcs into convertion_error .
    cl_demo_output=>write_text( convertion_error->if_message~get_longtext( ) ).
endtry.


* PART 4
* Save the file to the PC file system.
*
data: download_error type ref to cx_root.
try.
    call function 'GUI_DOWNLOAD'
      exporting
*       bin_filesize            = encoded_content_length
        filename                = 'C:\Users\r\Desktop\table_test_convert.xls' " change the path
        filetype                = 'BIN'
*       APPEND                  = ' '
*       WRITE_FIELD_SEPARATOR   = ' '
*       HEADER                  = '00'
*       TRUNC_TRAILING_BLANKS   = ' '
*       WRITE_LF                = 'X'
*       COL_SELECT              = ' '
*       COL_SELECT_MASK         = ' '
*       DAT_MODE                = ' '
*       CONFIRM_OVERWRITE       = ' '
*       NO_AUTH_CHECK           = ' '
*       CODEPAGE                = '4103'
*       IGNORE_CERR             = ABAP_TRUE
*       REPLACEMENT             = '#'
*       WRITE_BOM               = 'X'
*       TRUNC_TRAILING_BLANKS_EOL       = 'X'
*       WK1_N_FORMAT            = ' '
*       WK1_N_SIZE              = ' '
*       WK1_T_FORMAT            = ' '
*       WK1_T_SIZE              = ' '
*       WRITE_LF_AFTER_LAST_LINE        = ABAP_TRUE
*       SHOW_TRANSFER_STATUS    = ABAP_TRUE
*       VIRUS_SCAN_PROFILE      = '/SCET/GUI_DOWNLOAD'
* IMPORTING
*       FILELENGTH              =
      tables
        data_tab                = encoded_content
*       FIELDNAMES              =
      exceptions
        file_write_error        = 1
        no_batch                = 2
        gui_refuse_filetransfer = 3
        invalid_type            = 4
        no_authority            = 5
        unknown_error           = 6
        header_not_allowed      = 7
        separator_not_allowed   = 8
        filesize_not_allowed    = 9
        header_too_long         = 10
        dp_error_create         = 11
        dp_error_send           = 12
        dp_error_write          = 13
        unknown_dp_error        = 14
        access_denied           = 15
        dp_out_of_memory        = 16
        disk_full               = 17
        dp_timeout              = 18
        file_not_found          = 19
        dataprovider_exception  = 20
        control_flush_error     = 21
        others                  = 22.
    if sy-subrc <> 0.
      cl_demo_output=>write_data( sy-subrc ).
    endif.
  catch cx_root into download_error.
    cl_demo_output=>write_text( download_error->get_longtext( ) ).
endtry.

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
      ,lv_send         type ad_smtpadr value 'robinson.crusoe@island.com'
      ,lv_sent_to_all  type os_boolean
      .
"create send request
lo_send_request = cl_bcs=>create_persistent( ).

"create message body and subject
append 'Dear Friend,' to lt_message_body.
append initial line to lt_message_body.
append 'Please review the attached Microsoft Excel spreadsheet.' to lt_message_body.
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
        i_attachment_type = 'XLS'
        i_attachment_subject = 'table_test_convert'
        i_att_content_hex = encoded_content
        ).

  catch cx_document_bcs into lx_document_bcs.
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
"lo_recipient = cl_cam_address_bcs=>create_internet_address('friday@island.com' ). " set the e-mail address

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


cl_demo_output=>display( ).
