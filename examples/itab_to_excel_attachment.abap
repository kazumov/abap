class worker definition.
  public section.
    methods do
      importing some_tab type index table.
endclass.

class worker implementation.
  method do.
    constants: c_tab type c value cl_bcs_convert=>gc_tab,
               c_cr  type c value cl_bcs_convert=>gc_crlf.

    data: r_des           type ref to cl_abap_structdescr,
          r_some_tab_line type ref to data, " for header composition
          r_some_tab      type ref to data,
          r_exc           type ref to cx_root,
          t_comp          type abap_component_tab,
          line            type string, " value \t balue \t value...
          lines           type string, " line \n line \n line...
          cell            type string.

    " HEADER line composition
    try.
        create data r_some_tab_line like line of some_tab.
        r_des ?= cl_abap_typedescr=>describe_by_data_ref( r_some_tab_line ).
        t_comp = r_des->get_components( ).

        loop at t_comp assigning field-symbol(<f_comp>).
          concatenate line c_tab <f_comp>-name into line.
        endloop.

        replace regex '^\t' in line with ''. " delete the first exta tab symbol (x0900)

        lines = line. " the first line, heades, no \n at the end

      catch cx_root into r_exc.
        cl_demo_output=>write_text( line ).
        cl_demo_output=>write_data( some_tab ).
        cl_demo_output=>write_text( r_exc->get_longtext( ) ).
        cl_demo_output=>write_text( r_exc->get_text( ) ).
    endtry.


    " BODY
    clear: line.
    data: str_cast type string.
    field-symbols: <cell> type any.
    try.
*        create data r_some_tab like line of some_tab.
*        r_des ?= cl_abap_typedescr=>describe_by_data_ref( r_some_tab ).
*        t_comp = r_des->get_components( ).

        loop at some_tab assigning field-symbol(<f_some_tab>).
          loop at t_comp assigning <f_comp>.
            assign component <f_comp>-name of structure <f_some_tab> to <cell>.
            str_cast = <cell>.
            concatenate line c_tab str_cast into line.
            unassign <cell>.
            clear str_cast.
          endloop.
          replace regex '^\t' in line with ''. " delete the first extra tab symbol (x0900)
          concatenate lines c_cr line into lines. " add line and \n
          clear line.
        endloop.

      catch cx_root into r_exc.
        cl_demo_output=>write_text( line ).
        cl_demo_output=>write_data( some_tab ).
        cl_demo_output=>write_text( r_exc->get_longtext( ) ).
        cl_demo_output=>write_text( r_exc->get_text( ) ).
    endtry.

    data: gt_binary_content type solix_tab,
          gv_size type so_obj_len.

    try.
        cl_bcs_convert=>string_to_solix(
        exporting
        iv_string = lines
        iv_codepage = '4103'  "suitable for MS Excel, leave empty
        iv_add_bom = 'X'      "for other doc types
        importing
        et_solix = gt_binary_content
        ev_size = gv_size ).
      catch cx_bcs.
        " message e445(so).
    endtry.



    cl_demo_output=>write_text( lines ).
    cl_demo_output=>write_data( gt_binary_content ).
    cl_demo_output=>write_data( gv_size ).
    
* gt_binary_content is ready to be attached to e-mail    

  endmethod.
endclass.
