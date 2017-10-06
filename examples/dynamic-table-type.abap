*&---------------------------------------------------------------------*
*& Report Z_DIN_ITAB_TRANSF_TO_CLASS
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report z_din_itab_transf_to_class.

types: begin of rep,
         str1 type string,
         str2 type string,
         int3 type i,
       end of rep.
       
data: w_rep type rep,
      t_rep type table of rep.

w_rep-str1 = 'all'.
w_rep-str2 = 'you'.
w_rep-int3 = 1.
append w_rep to t_rep.

w_rep-str1 = 'need'.
w_rep-str2 = 'is'.
w_rep-int3 = 2.
append w_rep to t_rep.

w_rep-str1 = 'love'.
w_rep-str2 = '!'.
w_rep-int3 = 3.
append w_rep to t_rep.

class worker definition.
  public section.
    methods do
      importing some_tab type index table.
endclass.

class worker implementation.
  method do.
    constants: c_tab type c value cl_bcs_convert=>gc_tab,
               c_cr  type c value cl_bcs_convert=>gc_crlf.

    data: r_des      type ref to cl_abap_structdescr,
          r_some_tab type ref to data,
          r_exc      type ref to cx_root,
          t_comp     type abap_component_tab,
          line       type string,
          cell       type string.
    field-symbols: <f_some_tab> type any,
                   <f_comp>     like line of t_comp.
    try.
        create data r_some_tab like line of some_tab.
        assign r_some_tab->* to <f_some_tab>.

        r_des ?= cl_abap_typedescr=>describe_by_data( <f_some_tab> ).
        t_comp = r_des->get_components( ).

        loop at t_comp assigning <f_comp>.
          concatenate line c_tab <f_comp>-name into line.
        endloop.

        concatenate line c_cr into line.

      catch cx_root into r_exc.
        cl_demo_output=>write_text( line ).
        cl_demo_output=>write_data( some_tab ).
        cl_demo_output=>write_text( r_exc->get_longtext( ) ).
        cl_demo_output=>write_text( r_exc->get_text( ) ).
    endtry.
  endmethod.
endclass.


class processor definition.
  public section.
    methods work
      importing itab type index table.
endclass.

class processor implementation.
  method work.
    data r_w type ref to worker.
    create object r_w.
    field-symbols: <t> type any table.
    assign itab to <t>.
    r_w->do( some_tab = <t> ).
  endmethod.
endclass.


data: r_pr type ref to processor.

start-of-selection.

  create object r_pr.
  r_pr->work( itab = t_rep ).

  cl_demo_output=>display( ).
