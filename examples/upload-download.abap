data: it_pdf     type table of tline,
      ls_pdf     like line of it_pdf,
      lt_data1   type table of x255,
      lv_content type xstring,
      length     like sy-tabix.
field-symbols <fs_x> type x.
call function 'GUI_UPLOAD'
  exporting
    filename   = 'C:\Documents and Settings\downloaded.pdf '
    filetype   = 'BIN'
  importing
    filelength = length
  tables
    data_tab   = it_pdf.
loop at it_pdf into ls_pdf.
  assign ls_pdf to <fs_x> casting.
  concatenate lv_content <fs_x> into lv_content in byte mode.
endloop.

call function 'SCMS_XSTRING_TO_BINARY'
  exporting
    buffer        = lv_content
  importing
    output_length = length
  tables
    binary_tab    = lt_data1.
call function 'GUI_DOWNLOAD'
  exporting
    bin_filesize = length
    filename     = 'C:\Documents and Settings\test.pdf'
    filetype     = 'BIN'
  tables
    data_tab     = lt_data1.
