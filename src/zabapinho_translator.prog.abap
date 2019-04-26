REPORT zabapinho_translator MESSAGE-ID 00.
*----------------------------------------------------------------------*
* Created by Nuno Morais for objects translations
* Version 1.0
*----------------------------------------------------------------------*

*--------------------------------------------------------- GLOBAL DATA *
TABLES t002.                        "Language Keys
TYPE-POOLS: abap, icon, ole2, sscr. "Only for old versions
*----------------------------------------------------------- Constants *
CONSTANTS:
  gc_r3tr          TYPE pgmid        VALUE 'R3TR',    "Main object
  gc_temp          TYPE developclass VALUE '$TMP',    "Local development class
  gc_excel_ext     TYPE string       VALUE 'XLSX',    "Excel file extention
  gc_ieq(3)        TYPE c            VALUE 'IEQ',     "Ranges
  gc_pgmid(5)      TYPE c            VALUE 'PGMID',   "Fields names
  gc_object(6)     TYPE c            VALUE 'OBJECT',
  gc_lxe_object(3) TYPE c            VALUE 'LXE',
  gc_textkey(8)    TYPE c            VALUE 'TEXTPAIR',
  gc_length(6)     TYPE c            VALUE 'LENGTH',
  gc_desc(11)      TYPE c            VALUE 'DESCRIPTION'.
*---------------------------------------------------------- Structures *
*---------- Objects ----------*
TYPES:
  BEGIN OF gty_objects,
    status   TYPE icon_d,       "Check status
    pgmid    TYPE pgmid,        "Program ID in Requests and Tasks
    object   TYPE trobjtype,    "Object Type
    obj_name TYPE sobj_name,    "Object Name in Object Directory
    obj_desc TYPE ddtext,       "Object Explanatory short text
    slang    TYPE spras,        "Source Language
    tlangs   TYPE string,       "Target Languages
    stattrn  TYPE icon_d,       "Initial Translation status of an Object
    statproc TYPE icon_d,       "Process Translation status of an Object
    devclass TYPE developclass, "Development Package
    target   TYPE tr_target,    "Transport Target of Request
  END OF gty_objects.
*---------- LXE Object Lists ----------*
TYPES:
  BEGIN OF gty_colob,
    pgmid    TYPE pgmid,        "Program ID in Requests and Tasks
    object   TYPE trobjtype,    "Object Type
    obj_name TYPE sobj_name.    "Object Name in Object Directory
        INCLUDE TYPE lxe_colob. "Object Lists
TYPES:
  END OF gty_colob.
*---------- Languages Informations ----------*
TYPES:
  BEGIN OF gty_languages,
    r3_lang(2) TYPE c,          "R3 Language (Char 2)
    laiso      TYPE laiso,      "Language according to ISO 639
    o_language TYPE lxeisolang, "Translation Language
    text       TYPE sptxt,      "Name of Language
  END OF gty_languages.

SET EXTENDED CHECK OFF.
DATA:
  gt_objects    TYPE TABLE OF gty_objects,    "Objects to transport
  gt_objs_desc  TYPE TABLE OF ko100,          "Objects prograns IDs
  gt_objs_colob TYPE TABLE OF gty_colob,      "LXE Object Lists
  gt_languages  TYPE TABLE OF gty_languages.  "Target Languages Informations
*----------------------------------------------------------- Variables *
DATA:
  gv_percent  TYPE i,       "Progress bar percentage
  gv_tlangs   TYPE string,  "Target Languages
  gv_msg_text TYPE string,  "All Global Exceptions Text
  gv_object   TYPE trobjtype.
*------------------------------------------------------------- Objects *
DATA:
   go_objects TYPE REF TO cl_salv_table,  "Objects ALV
   go_exp     TYPE REF TO cx_root.        "Abstract Superclass for All Global Exceptions
SET EXTENDED CHECK ON.

*-------------------------------------- CLASS HANDLE EVENTS DEFINITION *
CLASS lcl_handle_events DEFINITION FINAL.
  PUBLIC SECTION.
    METHODS:
      on_user_command FOR EVENT added_function OF cl_salv_events
        IMPORTING e_salv_function,
      on_double_click FOR EVENT double_click OF cl_salv_events_table
        IMPORTING row column.                               "#EC NEEDED
ENDCLASS.                    "lcl_handle_events DEFINITION

*---------------------------------------------------- SELECTION SCREEN *
*---------------------------------------------------- Object selection *
SELECTION-SCREEN BEGIN OF BLOCK b01 WITH FRAME TITLE text-b01.
SELECTION-SCREEN SKIP 1.
*---------- Workbench object ----------*
PARAMETERS r_obj RADIOBUTTON GROUP rbt USER-COMMAND rbt DEFAULT 'X'.
SELECTION-SCREEN BEGIN OF LINE.
SELECTION-SCREEN POSITION 4.
PARAMETERS:
  p_pgmid  TYPE pgmid DEFAULT gc_r3tr,  "Program ID in Requests and Tasks
  p_object TYPE trobjtype,              "Object Type
  p_obj_n  TYPE sobj_name.              "Object Name in Object Directory
SELECTION-SCREEN END OF LINE.
*---------- Transport request ----------*
SELECTION-SCREEN SKIP 1.
PARAMETERS r_tr RADIOBUTTON GROUP rbt.
SELECTION-SCREEN BEGIN OF LINE.
SELECTION-SCREEN POSITION 4.
PARAMETERS p_tr TYPE trkorr.  "Transport request
SELECTION-SCREEN END OF LINE.
SELECTION-SCREEN SKIP 1.
SELECTION-SCREEN END OF BLOCK b01.
*-------------------------------------------------- Translation options *
SELECTION-SCREEN BEGIN OF BLOCK b02 WITH FRAME TITLE text-b02.
SELECTION-SCREEN SKIP 1.
PARAMETERS p_slang TYPE spras DEFAULT 'EN'.           "Source Language
SELECT-OPTIONS so_tlang FOR t002-spras NO INTERVALS.  "Target Languages
SELECTION-SCREEN SKIP 1.
PARAMETERS p_dep AS CHECKBOX DEFAULT abap_true.  "Dependencies check
PARAMETERS p_ow  AS CHECKBOX.  "Overwrite existent translations
SELECTION-SCREEN END OF BLOCK b02.

*--------------------------------------------- SELECTION SCREEN EVENTS *
*------------------------------------------------------- Program ID F4 *
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_pgmid.
  PERFORM pgmid_f4.
*------------------------------------------------------ Object Type F4 *
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_object.
  PERFORM object_f4.
*------------------------------------------------------ Object Name F4 *
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_obj_n.
  PERFORM object_name_f4.
*--------------------------------------------------- Transport request *
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_tr.
  CALL FUNCTION 'TR_F4_REQUESTS'
    IMPORTING
      ev_selected_request = p_tr.
*----------------------------------------- Selection Screen Events PAI *

AT SELECTION-SCREEN.
  PERFORM screen_pai.

*------------------------------------------------------- REPORT EVENTS *
*----------------------------------------------- Initialization events *
INITIALIZATION.
  PERFORM load_of_program.
*---------------------------------------------------- Executing events *
START-OF-SELECTION.
  PERFORM run_checks.

END-OF-SELECTION.
  PERFORM display_objects.

*--------------------------------------------------------------- FORMS *
*------------------------------------------------ Form LOAD_OF_PROGRAM *
FORM load_of_program.
  DATA:
    lt_restrict  TYPE sscr_restrict,  "Select Options Restrict
    lt_opt_list  TYPE sscr_opt_list,
    lt_associate TYPE sscr_ass.       "selection screen object

*---------- Fill Program IDs ----------*
  CALL FUNCTION 'TR_OBJECT_TABLE'
    TABLES
      wt_object_text = gt_objs_desc.

*---------- SC restrict SOs ----------*
  lt_opt_list-name = gc_ieq+1(2).
  lt_opt_list-options-eq = abap_true.
  APPEND lt_opt_list TO lt_restrict-opt_list_tab.
  lt_associate-kind    = 'S'.
  lt_associate-name    = 'SO_TLANG'.
  lt_associate-sg_main = gc_ieq(1).
  lt_associate-sg_addy = space.
  lt_associate-op_main = gc_ieq+1(2).
  lt_associate-op_addy = gc_ieq+1(2).
  APPEND lt_associate TO lt_restrict-ass_tab.

  CALL FUNCTION 'SELECT_OPTIONS_RESTRICT'
    EXPORTING
      restriction            = lt_restrict
    EXCEPTIONS
      too_late               = 1
      repeated               = 2
      selopt_without_options = 3
      selopt_without_signs   = 4
      invalid_sign           = 5
      empty_option_list      = 6
      invalid_kind           = 7
      repeated_kind_a        = 8
      OTHERS                 = 9.

  IF sy-subrc IS NOT INITIAL.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
ENDFORM.                    " LOAD_OF_PROGRAM
*------------------------------------------------------- Form PGMID_F4 *
FORM pgmid_f4 .
  DATA lt_pgmids TYPE TABLE OF ko101.  "Program IDs with Description

*---------- Read PGMID ----------*
  CALL FUNCTION 'TR_PGMID_TABLE'
    TABLES
      wt_pgmid_text = lt_pgmids.
*---------- Set PGMID F4 ----------*
  CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'              "#EC FB_RC
    EXPORTING
      retfield        = 'PGMID'
      dynpprog        = sy-cprog
      value_org       = 'S'
      dynpnr          = '1000'
      dynprofield     = 'TRE071X-PGMID'
    TABLES
      value_tab       = lt_pgmids
    EXCEPTIONS
      parameter_error = 1
      no_values_found = 2
      OTHERS          = 3.
ENDFORM.                    " PGMID_F4
*------------------------------------------------------ Form OBJECT_F4 *
FORM object_f4.
  DATA:
    lt_shlp          TYPE shlp_descr,               "Description of Search Help
    lt_return_values TYPE TABLE OF ddshretval,      "Interface Structure Search Help
    ls_return_values LIKE LINE OF lt_return_values,
    lv_rc            TYPE sysubrc.                  "Return Value of ABAP Statements
  FIELD-SYMBOLS <interface> TYPE ddshiface.         "Interface description of a F4 help method

*---------- Get search help ----------*
  CALL FUNCTION 'F4IF_GET_SHLP_DESCR'
    EXPORTING
      shlpname = 'SCTSOBJECT'
    IMPORTING
      shlp     = lt_shlp.
*---------- Fill search help ----------*
  LOOP AT lt_shlp-interface ASSIGNING <interface>.
    IF <interface>-shlpfield = gc_object.
      <interface>-valfield = abap_true.
      <interface>-value    = gv_object.
    ENDIF.
    IF <interface>-shlpfield = gc_pgmid.
      <interface>-valfield = abap_true.
      <interface>-value    = p_pgmid.
    ENDIF.
  ENDLOOP.
*---------- Call search help ----------*
  CALL FUNCTION 'F4IF_START_VALUE_REQUEST'
    EXPORTING
      shlp          = lt_shlp
    IMPORTING
      rc            = lv_rc
    TABLES
      return_values = lt_return_values.
*---------- Set search help return ----------*
  IF lv_rc IS INITIAL.
    READ TABLE lt_return_values INTO ls_return_values WITH KEY fieldname = gc_object.
    IF sy-subrc IS INITIAL.
      p_object = ls_return_values-fieldval.
    ENDIF.
    READ TABLE lt_return_values INTO ls_return_values WITH KEY fieldname = gc_pgmid.
    IF sy-subrc IS INITIAL.
      p_pgmid = ls_return_values-fieldval.
    ENDIF.
  ENDIF.
ENDFORM." OBJECT_F4
*------------------------------------------------- Form OBJECT_NAME_F4 *
FORM object_name_f4.
  DATA lv_object_type TYPE seu_obj.  "Object type

*---------- Get objects repository information ----------*
  lv_object_type = p_object.
  CALL FUNCTION 'REPOSITORY_INFO_SYSTEM_F4'                 "#EC FB_RC
    EXPORTING
      object_type          = lv_object_type
      object_name          = p_obj_n
    IMPORTING
      object_name_selected = p_obj_n
    EXCEPTIONS
      cancel               = 1
      wrong_type           = 2
      OTHERS               = 3.
ENDFORM.                    " OBJECT_NAME_F4
*----------------------------------------------------- Form SCREEN_PAI *
FORM screen_pai .
  gv_object = p_object.
*---------- Execute ----------*
  IF sy-ucomm = 'ONLI'.
    TRY.
*---------- Check and Read languages informations ----------*
        PERFORM check_languages.

        CASE abap_true.
*---------- Add object or dev class objects ----------*
          WHEN r_obj.
            PERFORM execute_add_objects.
*---------- Add TR objects ----------*
          WHEN r_tr.
            PERFORM execute_add_from_transport.
        ENDCASE.

*---------- Check if objects found ----------*
        IF gt_objects IS INITIAL.
          MESSAGE e398 WITH text-m03 space space space DISPLAY LIKE 'W'.  "Object not found
        ENDIF.
        IF p_ow IS NOT INITIAL.
          MESSAGE w398 WITH text-m08 space space space. "Overwrite existent translations activated
        ENDIF.

      CATCH cx_root INTO go_exp.                         "#EC CATCH_ALL
        gv_msg_text = go_exp->get_text( ).
        MESSAGE s398 WITH gv_msg_text space space space DISPLAY LIKE 'E'. "Critical error
    ENDTRY.
  ENDIF.
ENDFORM.                    " SCREEN_PAI

*--------------------------------------------------- Form PROGRESS_BAR *
FORM progress_bar USING i_value TYPE itex132 i_tabix TYPE i.
  DATA:
    lv_text(40),
    lv_percentage TYPE p,
    lv_percent_char(3).

  lv_percentage = ( i_tabix / 100 ) * 100.
  lv_percent_char = lv_percentage.
  SHIFT lv_percent_char LEFT DELETING LEADING space.
  CONCATENATE i_value '...' INTO i_value.
  CONCATENATE i_value lv_percent_char text-pb1 INTO lv_text SEPARATED BY space.

  IF lv_percentage GT gv_percent OR i_tabix = 1.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = lv_percentage
        text       = lv_text.

    gv_percent = lv_percentage.
  ENDIF.
ENDFORM.                    " PROGRESS_BAR
*------------------------------------------------ Form DISPLAY_OBJECTS *
FORM display_objects.
  DATA:
    lr_events        TYPE REF TO cl_salv_events_table,      "ALV Events
    lr_display       TYPE REF TO cl_salv_display_settings,  "ALV Output Appearance
    lr_columns       TYPE REF TO cl_salv_columns_table,     "ALV Columns
    lr_column        TYPE REF TO cl_salv_column_table,
    lr_selections    TYPE REF TO cl_salv_selections,        "ALV Selections
    lr_layout        TYPE REF TO cl_salv_layout,            "ALV Layout
    lo_event_handler TYPE REF TO lcl_handle_events.         "ALV Events Handler
  DATA:
    lt_column_ref TYPE salv_t_column_ref, "Columns of ALV List
    ls_column_ref TYPE salv_s_column_ref,
    ls_key        TYPE salv_s_layout_key. "Layout Key
  DATA:
    lv_title   TYPE lvc_title,  "ALV title
    lv_lines   TYPE i,          "Number of objects
    lv_lines_c TYPE string.

  PERFORM progress_bar USING text-p04 '90'. "Display objects
  IF go_objects IS NOT BOUND. "Create ALV
    TRY.
        IF lines( gt_objects ) = 1.
          MESSAGE s398 WITH text-m06 space space space DISPLAY LIKE 'W'.  "No dependecies found
        ELSE.
          SORT gt_objects BY pgmid object obj_name.
        ENDIF.
*---------- Create ALV ----------*
        cl_salv_table=>factory( IMPORTING r_salv_table = go_objects
                                 CHANGING t_table      = gt_objects ).
*---------- Set ALV Functions ----------*
        go_objects->set_screen_status(
          pfstatus      = 'STATUS'
          report        = sy-cprog
          set_functions = go_objects->c_functions_all ).
*---------- Set Layout ----------*
        lr_layout = go_objects->get_layout( ).
        ls_key-report = sy-repid.
        lr_layout->set_key( ls_key ).
        lr_layout->set_save_restriction( ).
*---------- Set ALV selections ----------*
        lr_selections = go_objects->get_selections( ).
        lr_selections->set_selection_mode( if_salv_c_selection_mode=>row_column ).
*---------- Set ALV Display and Title ----------*
        lr_display = go_objects->get_display_settings( ).
        lr_display->set_striped_pattern( if_salv_c_bool_sap=>true ).
        lv_lines = lines( gt_objects ).
        lv_lines_c = lv_lines.
        CONDENSE lv_lines_c NO-GAPS.
        CONCATENATE '(' lv_lines_c ')' INTO lv_lines_c.
        IF p_dep IS INITIAL.
          CONCATENATE text-t01 p_pgmid p_object p_obj_n text-t02 lv_lines_c INTO lv_title SEPARATED BY space.
        ELSE.
          CONCATENATE text-t01 p_pgmid p_object p_obj_n text-t02  text-t03 lv_lines_c INTO lv_title SEPARATED BY space.
        ENDIF.
        lr_display->set_list_header( lv_title ).
*---------- Set ALV Columns ----------*
        lr_columns = go_objects->get_columns( ).
        lr_columns->set_key_fixation( ).
        lr_columns->set_optimize( ).
        lt_column_ref = lr_columns->get( ).

        LOOP AT lt_column_ref INTO ls_column_ref. "Default format for all columns
          lr_column ?= lr_columns->get_column( ls_column_ref-columnname ).
          lr_column->set_f4( if_salv_c_bool_sap=>false ).
          lr_column->set_alignment( if_salv_c_alignment=>centered ).
*---------- Check status ----------*
          IF ls_column_ref-columnname = 'STATUS'.
            lr_column->set_key( if_salv_c_bool_sap=>true ).
            lr_column->set_short_text( text-c06 ).  "Check St.
            lr_column->set_medium_text( text-c06 ).
            lr_column->set_long_text( text-c06 ).
          ENDIF.
*---------- Object Keys ----------*
          IF ls_column_ref-columnname = gc_pgmid OR
             ls_column_ref-columnname = gc_object OR
             ls_column_ref-columnname = 'OBJ_NAME'.
            lr_column->set_key( if_salv_c_bool_sap=>true ).
          ENDIF.
*---------- Object name and development package ----------*
          IF ls_column_ref-columnname = 'OBJ_NAME' OR
             ls_column_ref-columnname = 'DEVCLASS'.
            lr_column->set_alignment( if_salv_c_alignment=>left ).
          ENDIF.
*---------- Object description ----------*
          IF ls_column_ref-columnname = 'OBJ_DESC'.
            lr_column->set_alignment( if_salv_c_alignment=>left ).
            lr_column->set_short_text( text-c01 ).  "Obj. Desc.
            lr_column->set_medium_text( text-c01 ).
            lr_column->set_long_text( text-c01 ).
          ENDIF.
*---------- Source Language ----------*
          IF ls_column_ref-columnname = 'SLANG'.
            lr_column->set_short_text( text-c02 ).  "Source
            lr_column->set_medium_text( text-c02 ).
            lr_column->set_long_text( text-c02 ).
          ENDIF.
*---------- Target Languages ----------*
          IF ls_column_ref-columnname = 'TLANGS'.
            lr_column->set_short_text( text-c03 ).  "Targets
            lr_column->set_medium_text( text-c03 ).
            lr_column->set_long_text( text-c03 ).
          ENDIF.
*---------- Translation status ----------*
          IF ls_column_ref-columnname = 'STATTRN'.
            lr_column->set_short_text( text-c04 ).  "Trans. St.
            lr_column->set_medium_text( text-c04 ).
            lr_column->set_long_text( text-c04 ).
          ENDIF.
*---------- Process status ----------*
          IF ls_column_ref-columnname = 'STATPROC'.
            lr_column->set_short_text( text-c05 ).  "Proc. St.
            lr_column->set_medium_text( text-c05 ).
            lr_column->set_long_text( text-c05 ).
          ENDIF.
        ENDLOOP.
*---------- Register ALV Events ----------*
        lr_events = go_objects->get_event( ).
        CREATE OBJECT lo_event_handler.
        SET HANDLER lo_event_handler->on_user_command FOR lr_events.
        SET HANDLER lo_event_handler->on_double_click FOR lr_events.
*---------- Display Objects ALV ----------*
        go_objects->display( ).

      CATCH cx_root INTO go_exp.                         "#EC CATCH_ALL
        gv_msg_text = go_exp->get_text( ).
        MESSAGE s398 WITH gv_msg_text space space space DISPLAY LIKE 'E'. "Critical error
    ENDTRY.

  ELSE. "Refresh ALV
    go_objects->refresh( ).
  ENDIF.
ENDFORM.                    " DISPLAY_OBJECTS
*---------------------------------------------------------- FORMS ADDS *

*&---------------------------------------------------------------------*
*&      Form  CHECK_LANGUAGES
*&---------------------------------------------------------------------*
FORM check_languages .
  DATA:
    ls_language   LIKE LINE OF gt_languages, "Languages Informations
    ls_tlang      LIKE LINE OF so_tlang.      "Target Languages

  IF p_slang IS INITIAL OR so_tlang[] IS INITIAL.
    MESSAGE e398 WITH text-m01 space space space DISPLAY LIKE 'W'.  "Please fill all required fields
  ENDIF.

  ls_tlang     = gc_ieq.
  ls_tlang-low = p_slang.
  APPEND ls_tlang TO so_tlang.
  REFRESH gt_languages.
  CLEAR gv_tlangs.

  LOOP AT so_tlang INTO ls_tlang WHERE low IS NOT INITIAL.
    CALL FUNCTION 'LXE_T002_CHECK_LANGUAGE'   "#EC FB_RC  "#EC CI_SUBRC
      EXPORTING
        r3_lang            = ls_tlang-low
      IMPORTING
        text               = ls_language-text       "Name of Language
        o_language         = ls_language-o_language "Translation Language
      EXCEPTIONS
        language_not_in_cp = 1
        unknown            = 2
        OTHERS             = 3.

    CALL FUNCTION 'CONVERSION_EXIT_ISOLA_OUTPUT'
      EXPORTING
        input  = ls_tlang-low
      IMPORTING
        output = ls_language-r3_lang.

    ls_language-laiso = ls_tlang-low. "Language according to ISO 639
    IF ls_tlang-low <> p_slang.
      IF gv_tlangs IS INITIAL.
        gv_tlangs = ls_language-r3_lang.
      ELSE.
        CONCATENATE gv_tlangs ls_language-r3_lang INTO gv_tlangs SEPARATED BY space.
      ENDIF.
    ENDIF.
    APPEND ls_language TO gt_languages.
    CLEAR ls_language.
  ENDLOOP.
  DELETE so_tlang WHERE low = p_slang.
ENDFORM.                    " CHECK_LANGUAGES

*&---------------------------------------------------------------------*
*&      Form  EXECUTE_ADD_OBJECTS
*&---------------------------------------------------------------------*
FORM execute_add_objects .
  DATA:
    lt_objectlist TYPE TABLE OF rseui_set,    "Transfer table (object list) - info system
    ls_objectlist LIKE LINE OF lt_objectlist,
    ls_env_dummy  TYPE senvi.                 "Object in Development Environment
  DATA lv_devclass TYPE devclass. "Package

  IF p_pgmid IS INITIAL OR p_object IS INITIAL OR p_obj_n IS INITIAL.
    MESSAGE e398 WITH text-m01 space space space DISPLAY LIKE 'W'.  "Please fill all required fields
  ENDIF.

  PERFORM progress_bar USING text-p01 '10'. "Adding object
  CASE p_object.
    WHEN 'DEVC'.  "Get all development package objects
      lv_devclass = p_obj_n.
      CALL FUNCTION 'RS_GET_OBJECTS_OF_DEVCLASS'            "#EC FB_RC
        EXPORTING
          devclass            = lv_devclass
        TABLES
          objectlist          = lt_objectlist
        EXCEPTIONS
          no_objects_selected = 1
          OTHERS              = 2.

      LOOP AT lt_objectlist INTO ls_objectlist.
        PERFORM check_add_object USING gc_r3tr ls_objectlist-obj_type ls_objectlist-obj_name ls_env_dummy.
      ENDLOOP.

    WHEN OTHERS.  "Add object
      PERFORM check_add_object USING p_pgmid p_object p_obj_n ls_env_dummy.
  ENDCASE.
ENDFORM.                    " EXECUTE_ADD_OBJECTS
*&---------------------------------------------------------------------*
*&      Form  EXECUTE_ADD_FROM_TRANSPORT
*&---------------------------------------------------------------------*
FORM execute_add_from_transport .

  DATA:
    lt_request_headers TYPE trwbo_request_headers,
    ls_request_headers TYPE trwbo_request_header,
    lt_objects         TYPE tr_objects,
    ls_object          TYPE e071,
    ls_env_dummy  TYPE senvi.

  IF p_tr IS INITIAL.
    MESSAGE e398 WITH text-m01 space space space DISPLAY LIKE 'W'.  "Please fill all required fields
  ENDIF.

  PERFORM progress_bar USING text-p01 '10'. "Adding object

*---------- Read Requests and Tasks ----------*
  CALL FUNCTION 'TR_READ_REQUEST_WITH_TASKS'
    EXPORTING
      iv_trkorr          = p_tr
    IMPORTING
      et_request_headers = lt_request_headers
    EXCEPTIONS
      invalid_input      = 1
      OTHERS             = 2.

  IF sy-subrc IS NOT INITIAL.
    MESSAGE e398 WITH text-m14 space space space DISPLAY LIKE 'W'.  "Error treating transport request
  ENDIF.
*---------- Read objects inside main request ----------*
  READ TABLE lt_request_headers INTO ls_request_headers WITH KEY trfunction = 'K'.
  IF sy-subrc IS NOT INITIAL.
    MESSAGE e398 WITH text-m14 space space space DISPLAY LIKE 'W'.  "Error treating transport request
  ENDIF.

  CALL FUNCTION 'TR_GET_OBJECTS_OF_REQ_AN_TASKS'
    EXPORTING
      is_request_header      = ls_request_headers
      iv_condense_objectlist = 'X'
    IMPORTING
      et_objects             = lt_objects
    EXCEPTIONS
      invalid_input          = 1
      OTHERS                 = 2.

  IF sy-subrc IS NOT INITIAL.
    MESSAGE e398 WITH text-m14 space space space DISPLAY LIKE 'W'.  "Error treating transport request
  ENDIF.

  CALL FUNCTION 'TR_SORT_OBJECT_AND_KEY_LIST'
    CHANGING
      ct_objects = lt_objects.

  LOOP AT lt_objects INTO ls_object.  "Add found objects to processing
    PERFORM check_add_object USING ls_object-pgmid ls_object-object ls_object-obj_name ls_env_dummy.
  ENDLOOP.

ENDFORM.                    " EXECUTE_ADD_FROM_TRANSPORT
*----------------------------------------------- Form CHECK_ADD_OBJECT *
FORM check_add_object USING value(i_pgmid) TYPE pgmid i_object TYPE any i_obj_n TYPE any is_env_tab TYPE senvi.
  DATA lo_wb_object TYPE REF TO cl_wb_object.  "Repository Object
  DATA:
    ls_tadir          TYPE tadir,               "Directory of Repository Objects
    ls_wb_object_type TYPE wbobjtype,           "Global WB Type
    ls_object         LIKE LINE OF gt_objects.  "Objects to transport line
  DATA:
    lv_tr_object   TYPE trobjtype,  "Object Type
    lv_tr_obj_name TYPE trobj_name, "Object Name in Object List
    lv_trans_pgmid TYPE pgmid.      "Program ID in Requests and Tasks

*---------- Object convertions ----------*
  IF i_pgmid <> gc_r3tr.
    SELECT pgmid UP TO 1 ROWS FROM tadir                "#EC CI_GENBUFF
      INTO i_pgmid
     WHERE object   = i_object
       AND obj_name = i_obj_n.
    ENDSELECT.
*---------- Is not a TADIR object and Conversion required ----------*
    IF sy-subrc IS NOT INITIAL.
      lv_tr_object   = i_object.
      lv_tr_obj_name = i_obj_n.
      cl_wb_object=>create_from_transport_key( EXPORTING p_object                = lv_tr_object
                                                         p_obj_name              = lv_tr_obj_name
                                               RECEIVING p_wb_object             = lo_wb_object
                                              EXCEPTIONS objecttype_not_existing = 1
                                                         empty_object_key        = 2
                                                         key_not_available       = 3
                                                         OTHERS                  = 4 ).
      IF sy-subrc IS INITIAL.
        lo_wb_object->get_global_wb_key( IMPORTING p_object_type     = ls_wb_object_type
                                        EXCEPTIONS key_not_available = 1
                                                   OTHERS            = 2 ).
        IF sy-subrc IS INITIAL.
          lo_wb_object->get_transport_key( IMPORTING p_pgmid           = lv_trans_pgmid "#EC CI_SUBRC
                                          EXCEPTIONS key_not_available = 1
                                                     OTHERS            = 2 ).
*---------- Check Program ID ----------*
          CASE lv_trans_pgmid.
            WHEN gc_r3tr.  "Main objects
              i_pgmid = lv_trans_pgmid.

            WHEN 'LIMU'.  "Sub object
              CALL FUNCTION 'GET_R3TR_OBJECT_FROM_LIMU_OBJ'
                EXPORTING
                  p_limu_objtype = lv_tr_object
                  p_limu_objname = lv_tr_obj_name
                IMPORTING
                  p_r3tr_objtype = lv_tr_object
                  p_r3tr_objname = lv_tr_obj_name
                EXCEPTIONS
                  no_mapping     = 1
                  OTHERS         = 2.
              IF sy-subrc IS INITIAL.
                ls_object-pgmid    = gc_r3tr.
                ls_object-object   = lv_tr_object.
                ls_object-obj_name = lv_tr_obj_name.
                PERFORM add_object USING ls_object.
                RETURN.
              ENDIF.

            WHEN OTHERS.  "Include objects
              i_pgmid = gc_r3tr.
              CALL FUNCTION 'GET_TADIR_TYPE_FROM_WB_TYPE'
                EXPORTING
                  wb_objtype        = ls_wb_object_type-subtype_wb
                IMPORTING
                  transport_objtype = lv_tr_object
                EXCEPTIONS
                  no_mapping_found  = 1
                  no_unique_mapping = 2
                  OTHERS            = 3.

              IF sy-subrc IS INITIAL.
                i_object = lv_tr_object.
                IF is_env_tab-encl_obj IS NOT INITIAL.
                  i_obj_n = is_env_tab-encl_obj.
                ENDIF.
              ENDIF.
          ENDCASE.
        ENDIF.  "Global WB key
      ENDIF.  "Transport_key
    ENDIF.  "No a TADIR
  ENDIF.  "Convertions

*---------- Check in TADIR ----------*
  SELECT SINGLE * FROM tadir
    INTO ls_tadir
   WHERE pgmid    = i_pgmid
     AND object   = i_object
     AND obj_name = i_obj_n.

*---------- Add object ----------*
  IF ls_tadir IS NOT INITIAL.
    MOVE-CORRESPONDING ls_tadir TO ls_object.
*---------- Set SAP Generated object status ----------*
    IF ls_tadir-genflag IS NOT INITIAL.
      ls_object-status = icon_led_yellow.
    ENDIF.
*---------- Add object to be checked ----------*
    PERFORM add_object USING ls_object.
*---------- Error Object not valid ----------*
  ELSE.
    IF lines( gt_objects ) > 0. "Skip first object
      ls_object-pgmid    = i_pgmid.
      ls_object-object   = i_object.
      ls_object-obj_name = i_obj_n.
      ls_object-status   = icon_led_red.
      PERFORM add_object USING ls_object.
    ENDIF.
  ENDIF.
ENDFORM.                    "check_add_object
*----------------------------------------------------- Form ADD_OBJECT *
FORM add_object USING ps_object TYPE gty_objects.
  DATA:
    ls_objs_desc LIKE LINE OF gt_objs_desc,  "Objects prograns ID line"Info Environment
    lt_devclass  TYPE scts_devclass,         "Development Packages
    ls_devclass  TYPE trdevclass.
  DATA:
    lv_object    TYPE trobjtype,  "Object Type
    lv_objname   TYPE sobj_name,  "Object Name in Object Directory
    lv_namespace TYPE namespace.  "Object Namespace

*---------- Check if already added ----------*
  READ TABLE gt_objects TRANSPORTING NO FIELDS WITH KEY pgmid    = ps_object-pgmid
                                                        object   = ps_object-object
                                                        obj_name = ps_object-obj_name.
  IF sy-subrc IS NOT INITIAL. "New object
*---------- Check if is customer object ----------*
    lv_object  = ps_object-object.
    lv_objname = ps_object-obj_name.
    CALL FUNCTION 'TRINT_GET_NAMESPACE'                     "#EC FB_RC
      EXPORTING
        iv_pgmid            = ps_object-pgmid
        iv_object           = lv_object
        iv_obj_name         = lv_objname
      IMPORTING
        ev_namespace        = lv_namespace
      EXCEPTIONS
        invalid_prefix      = 1
        invalid_object_type = 2
        OTHERS              = 3.

    IF lv_namespace = '/0CUST/'.  "Is customer object
*---------- Read object description ----------*
      READ TABLE gt_objs_desc INTO ls_objs_desc WITH KEY object = ps_object-object.
      IF sy-subrc IS INITIAL.
        ps_object-obj_desc = ls_objs_desc-text.  "Object type description
      ENDIF.
*---------- Read development class tecnical information ----------*
      IF ps_object-devclass IS INITIAL.
        SELECT SINGLE devclass FROM tadir
          INTO ps_object-devclass
         WHERE pgmid    = ps_object-pgmid
           AND object   = ps_object-object
           AND obj_name = ps_object-obj_name.
      ENDIF.

      IF ps_object-devclass IS NOT INITIAL AND ps_object-devclass <> gc_temp.
        ls_devclass-devclass = ps_object-devclass.
        APPEND ls_devclass TO lt_devclass.
        CALL FUNCTION 'TR_READ_DEVCLASSES'
          EXPORTING
            it_devclass = lt_devclass
          IMPORTING
            et_devclass = lt_devclass.
        READ TABLE lt_devclass INTO ls_devclass INDEX 1.
        IF sy-subrc IS INITIAL.
          ps_object-target = ls_devclass-target.  "Development package target
        ENDIF.
      ENDIF.

      ps_object-slang  = p_slang.
      ps_object-tlangs = gv_tlangs.
*---------- Add object to transport ----------*
      APPEND ps_object TO gt_objects.
    ENDIF.
  ENDIF.
ENDFORM.                    " ADD_OBJECT
*-------------------------------------------------------- FORMS CHECKS *
*----------------------------------------------------- Form RUN_CHECKS *
FORM run_checks .
  TRY.
*---------- Dependecies check ----------*
      PERFORM objects_dependencies_check.
*---------- Translations check ----------*
      PERFORM objects_translations_check.

    CATCH cx_root INTO go_exp.                           "#EC CATCH_ALL
      gv_msg_text = go_exp->get_text( ).
      MESSAGE s398 WITH gv_msg_text space space space DISPLAY LIKE 'E'. "Critical error
  ENDTRY.
ENDFORM.                    " RUN_CHECKS
*------------------------------------- Form OBJECTS_DEPENDENCIES_CHECK *
FORM objects_dependencies_check .
  DATA:
    lv_obj_type TYPE seu_obj,         "Object type
    lt_env_tab  TYPE TABLE OF senvi,  "Object to check dependencies
    ls_env_tab  TYPE senvi.           "Info Environment
  FIELD-SYMBOLS <ls_object> LIKE LINE OF gt_objects.  "Objects to transport

  PERFORM progress_bar USING text-p02 '30'. "Checking Dependecies
  LOOP AT gt_objects ASSIGNING <ls_object> WHERE status IS INITIAL.
*---------- Get object dependecies ----------*
    IF p_dep IS NOT INITIAL.
      REFRESH lt_env_tab.
      lv_obj_type = <ls_object>-object.
      CALL FUNCTION 'REPOSITORY_ENVIRONMENT_RFC'
        EXPORTING
          obj_type        = lv_obj_type
          object_name     = <ls_object>-obj_name
        TABLES
          environment_tab = lt_env_tab.

      DELETE lt_env_tab INDEX 1.  "Delete first line

*---------- Add founded dependecies ----------*
      LOOP AT lt_env_tab INTO ls_env_tab.                "#EC CI_NESTED
        PERFORM check_add_object USING space ls_env_tab-type ls_env_tab-object ls_env_tab.
      ENDLOOP.
    ENDIF.
    <ls_object>-status = icon_led_green.  "Status checked
  ENDLOOP.
ENDFORM.                    " OBJECTS_DEPENDENCIES_CHECK
*------------------------------------- Form OBJECTS_TRANSLATIONS_CHECK *
FORM objects_translations_check .
  DATA:
    lt_colob      TYPE TABLE OF lxe_colob,    "Object Lists
    ls_colob      LIKE LINE OF lt_colob,
    ls_objs_colob LIKE LINE OF gt_objs_colob, "LXE Objects
    ls_tlang      LIKE LINE OF so_tlang,      "Target Languages
    ls_language   LIKE LINE OF gt_languages,  "Languages Informations
    ls_slanguage  LIKE LINE OF gt_languages.
  DATA:
    lv_tr_obj_name TYPE trobj_name, "Object Name in Object List
    lv_stattrn     TYPE lxestattrn. "Translation Status of an Object
  FIELD-SYMBOLS <ls_object> LIKE LINE OF gt_objects.  "Objects to transport

  PERFORM progress_bar USING text-p03 '60'. "Checking Translations
*---------- Checking Translations ----------*
  REFRESH gt_objs_colob.
  LOOP AT gt_objects ASSIGNING <ls_object> WHERE status = icon_led_green.
    REFRESH lt_colob.
    lv_tr_obj_name = <ls_object>-obj_name.
    CALL FUNCTION 'LXE_OBJ_EXPAND_TRANSPORT_OBJ'
      EXPORTING
        pgmid           = <ls_object>-pgmid
        object          = <ls_object>-object
        obj_name        = lv_tr_obj_name
      TABLES
        ex_colob        = lt_colob
      EXCEPTIONS
        unknown_object  = 1
        unknown_ta_type = 2
        OTHERS          = 3.
*---------- Check Status ----------*
    IF sy-subrc IS NOT INITIAL. "Error
      <ls_object>-status = icon_led_red.
      CONTINUE.
    ENDIF.
    IF lt_colob IS INITIAL. "No translation
      <ls_object>-status = icon_led_yellow.
    ENDIF.

*---------- Initial Translation status ----------*
*---------- Add to global LXE object tables ----------*
    LOOP AT lt_colob INTO ls_colob.                      "#EC CI_NESTED
      MOVE-CORRESPONDING <ls_object> TO ls_objs_colob.
      MOVE-CORRESPONDING ls_colob    TO ls_objs_colob.
      APPEND ls_objs_colob TO gt_objs_colob.
      CLEAR ls_objs_colob.

      IF <ls_object>-stattrn <> icon_led_yellow.
*---------- Loop selected target language ----------*
        LOOP AT so_tlang INTO ls_tlang.                  "#EC CI_NESTED
          READ TABLE gt_languages INTO ls_language WITH KEY laiso = ls_tlang-low. "Read language tecnical info
          READ TABLE gt_languages INTO ls_slanguage WITH KEY laiso = p_slang.
          CLEAR lv_stattrn.
          CALL FUNCTION 'LXE_OBJ_TRANSLATION_STATUS2'
            EXPORTING
              t_lang  = ls_language-o_language
              s_lang  = ls_slanguage-o_language
              custmnr = ls_colob-custmnr
              objtype = ls_colob-objtype
              objname = ls_colob-objname
            IMPORTING
              stattrn = lv_stattrn.
          IF lv_stattrn = 'T'.  "Translated
            <ls_object>-stattrn = icon_led_green.
          ELSE.
            <ls_object>-stattrn = icon_led_yellow.
            EXIT.
          ENDIF.
        ENDLOOP.  "Target Languagues
      ENDIF.  "Status

    ENDLOOP.  "LXE Objects
  ENDLOOP.  "Objects
ENDFORM.                    " OBJECTS_TRANSLATIONS_CHECK
*------------------------------------------------------- FORMS OPTIONS *
*----------------------------------------------- Form CREATE_TRANSPORT *
FORM create_transport TABLES lt_objects TYPE STANDARD TABLE. "#EC PF_NO_TYPE
  DATA:
    lt_e071_temp  TYPE TABLE OF e071,         "Change & Transport System: Object Entries of Requests/Tasks
    lt_e071       TYPE TABLE OF e071,
    lt_e071k_temp TYPE TABLE OF e071k,
    ls_object     LIKE LINE OF gt_objects,
    lt_targets    TYPE TABLE OF tr_target,    "Transport Target of Request
    ls_target     LIKE LINE OF lt_targets,
    ls_objs_colob LIKE LINE OF gt_objs_colob, "LXE Objects
    ls_tlang      LIKE LINE OF so_tlang.      "Target Languages
  DATA:
    lv_order TYPE trkorr, "Request/Task
    lv_task  TYPE trkorr.

*---------- Check selected objects to transport ----------*
  LOOP AT lt_objects INTO ls_object.
    IF ls_object-devclass = gc_temp.
      MESSAGE i398 WITH text-m10 space space space DISPLAY LIKE 'E'.  "Request canceled, at least one object $TEMP detected
      RETURN.
    ENDIF.
    ls_target = ls_object-target.
    APPEND ls_target TO lt_targets.
  ENDLOOP.
*---------- Check targets ----------*
  SORT lt_targets.
  DELETE ADJACENT DUPLICATES FROM lt_targets.
  IF lines( lt_targets ) > 1. "Only one valid target
    MESSAGE i398 WITH text-m05 space space space. "Transport not allowed for multiple targets
    RETURN.
  ENDIF.

*---------- Add translations to transport ----------*
  LOOP AT lt_objects INTO ls_object.
    LOOP AT gt_objs_colob INTO ls_objs_colob WHERE pgmid    = ls_object-pgmid "#EC CI_NESTED
                                               AND object   = ls_object-object
                                               AND obj_name = ls_object-obj_name.
      LOOP AT so_tlang INTO ls_tlang.                    "#EC CI_NESTED
        CALL FUNCTION 'LXE_OBJ_CREATE_TRANSPORT_SE63'
          EXPORTING
            language = ls_tlang-low
            custmnr  = ls_objs_colob-custmnr
            objtype  = ls_objs_colob-objtype
            objname  = ls_objs_colob-objname
          TABLES
            ex_e071  = lt_e071_temp.
        APPEND LINES OF lt_e071_temp TO lt_e071.
        REFRESH lt_e071_temp.
      ENDLOOP.
    ENDLOOP.
  ENDLOOP.
*---------- Check selected translations ----------*
  IF lt_e071 IS INITIAL.
    MESSAGE i398 WITH text-m04 space space space. "No objects selected
    RETURN.
  ENDIF.
*---------- Create or Select transport request ----------*
  READ TABLE lt_targets INTO ls_target INDEX 1.
  CALL FUNCTION 'TRINT_ORDER_CHOICE'
    EXPORTING
      iv_tarsystem           = ls_target
    IMPORTING
      we_order               = lv_order
      we_task                = lv_task
    TABLES
      wt_e071                = lt_e071_temp
      wt_e071k               = lt_e071k_temp
    EXCEPTIONS
      no_correction_selected = 1
      display_mode           = 2
      object_append_error    = 3
      recursive_call         = 4
      wrong_order_type       = 5
      OTHERS                 = 6.
*---------- Add object to transport request ----------*
  IF sy-subrc IS INITIAL AND lv_task IS NOT INITIAL.
    REFRESH lt_e071k_temp.
    CALL FUNCTION 'TRINT_APPEND_COMM'
      EXPORTING
        wi_exclusive       = abap_false
        wi_sel_e071        = abap_true
        wi_sel_e071k       = abap_true
        wi_trkorr          = lv_task
      TABLES
        wt_e071            = lt_e071
        wt_e071k           = lt_e071k_temp
      EXCEPTIONS
        e071k_append_error = 1
        e071_append_error  = 2
        trkorr_empty       = 3
        OTHERS             = 4.
*---------- Sort and compress request --------*
    IF sy-subrc IS INITIAL. "Added with sucess
      CALL FUNCTION 'TR_SORT_AND_COMPRESS_COMM' "#EC FB_RC   "#EC CI_SUBRC
        EXPORTING
          iv_trkorr                      = lv_task
        EXCEPTIONS
          trkorr_not_found               = 1
          order_released                 = 2
          error_while_modifying_obj_list = 3
          tr_enqueue_failed              = 4
          no_authorization               = 5
          OTHERS                         = 6.
      MESSAGE i398 WITH text-m07 lv_order space space DISPLAY LIKE 'S'.  "Objects added to request
    ELSE.
      MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'.  "Executed with errors
    ENDIF.  "Added

  ELSE.
    MESSAGE s398 WITH text-m09 space space space DISPLAY LIKE 'W'.  "Transport canceled
  ENDIF.
ENDFORM.                    " CREATE_TRANSPORT
*---------------------------------------------- Form COPY_TRANSLATIONS *
FORM copy_translations TABLES lt_objects TYPE STANDARD TABLE. "#EC PF_NO_TYPE
  DATA:
      ls_object     LIKE LINE OF gt_objects,    "Objects to transport
      ls_objs_colob LIKE LINE OF gt_objs_colob, "LXE Objects
      ls_tlang      LIKE LINE OF so_tlang,      "Target Languages
      lt_pcx_s1     TYPE TABLE OF lxe_pcx_s1,   "Text Pairs
      ls_pcx_s1     LIKE LINE OF lt_pcx_s1,
      ls_tlanguage   LIKE LINE OF gt_languages, "Languages Informations
      ls_slanguage  LIKE LINE OF gt_languages.
  DATA:
    lv_pstatus TYPE lxestatprc, "Process Status
    lv_stattrn TYPE lxestattrn. "Translation Status of an Object
  FIELD-SYMBOLS:
    <ls_pcx_s1> LIKE LINE OF lt_pcx_s1,   "Text Pairs
    <ls_object> LIKE LINE OF gt_objects.  "Objects to translate

*---------- Read source language details ----------*
  READ TABLE gt_languages INTO ls_slanguage WITH KEY laiso = p_slang.

*---------- Loop all selected objects ----------*
  LOOP AT lt_objects INTO ls_object.

*---------- Loop LXE Sub-Objects ----------*
    LOOP AT gt_objs_colob INTO ls_objs_colob WHERE pgmid    = ls_object-pgmid "#EC CI_NESTED
                                               AND object   = ls_object-object
                                               AND obj_name = ls_object-obj_name.

*---------- Loop selected target language ----------*
      LOOP AT so_tlang INTO ls_tlang.                    "#EC CI_NESTED
*---------- Read target language details ----------*
        CLEAR ls_tlanguage.
        READ TABLE gt_languages INTO ls_tlanguage WITH KEY laiso = ls_tlang-low. "Read language tecnical info
*---------- Read texts ----------*
        CLEAR lv_pstatus.
        REFRESH lt_pcx_s1.
        CALL FUNCTION 'LXE_OBJ_TEXT_PAIR_READ'
          EXPORTING
            t_lang    = ls_tlanguage-o_language
            s_lang    = ls_slanguage-o_language
            custmnr   = ls_objs_colob-custmnr
            objtype   = ls_objs_colob-objtype
            objname   = ls_objs_colob-objname
            read_only = space
          IMPORTING
            pstatus   = lv_pstatus
          TABLES
            lt_pcx_s1 = lt_pcx_s1.
        IF lv_pstatus <> 'S' OR lt_pcx_s1 IS INITIAL. "Not Successful or empty
          CONTINUE.
        ENDIF.
*---------- Copy and check Overwrite ----------*
        IF p_ow IS INITIAL.
          LOOP AT lt_pcx_s1 ASSIGNING <ls_pcx_s1> WHERE t_text IS INITIAL. "#EC CI_NESTED
            <ls_pcx_s1>-t_text = <ls_pcx_s1>-s_text.
          ENDLOOP.
        ELSE.
          LOOP AT lt_pcx_s1 ASSIGNING <ls_pcx_s1>.       "#EC CI_NESTED
            <ls_pcx_s1>-t_text = <ls_pcx_s1>-s_text.
          ENDLOOP.
        ENDIF.
*---------- Update texts ----------*
        CALL FUNCTION 'LXE_OBJ_TEXT_PAIR_WRITE'
          EXPORTING
            t_lang    = ls_tlanguage-o_language
            s_lang    = ls_slanguage-o_language
            custmnr   = ls_objs_colob-custmnr
            objtype   = ls_objs_colob-objtype
            objname   = ls_objs_colob-objname
          TABLES
            lt_pcx_s1 = lt_pcx_s1.
*---------- Create proposal and check status ----------*
        LOOP AT lt_pcx_s1 INTO ls_pcx_s1 WHERE t_text IS NOT INITIAL. "#EC CI_NESTED
          CALL FUNCTION 'LXE_PP1_PROPOSAL_EDIT_SE63'
            EXPORTING
              t_lang         = ls_tlanguage-o_language
              s_lang         = ls_slanguage-o_language
              custmnr        = ls_objs_colob-custmnr
              objtype        = ls_objs_colob-objtype
              domatyp        = ls_objs_colob-domatyp
              domanam        = ls_objs_colob-domanam
              pcx_s1         = ls_pcx_s1
              direct_command = 'ASTX'
              direct_status  = '69'.
        ENDLOOP.  "Proposal

*---------- Check and update translation status ----------*
        READ TABLE gt_objects ASSIGNING <ls_object> WITH KEY pgmid    = ls_object-pgmid
                                                             object   = ls_object-object
                                                             obj_name = ls_object-obj_name.
        IF sy-subrc IS INITIAL AND <ls_object>-statproc <> icon_led_yellow.
          CLEAR lv_stattrn.
          CALL FUNCTION 'LXE_OBJ_TRANSLATION_STATUS2'
            EXPORTING
              t_lang  = ls_tlanguage-o_language
              s_lang  = ls_slanguage-o_language
              custmnr = ls_objs_colob-custmnr
              objtype = ls_objs_colob-objtype
              objname = ls_objs_colob-objname
            IMPORTING
              stattrn = lv_stattrn.
*---------- Process Translation status ----------*
          IF lv_stattrn = 'T'.  "Translated
            <ls_object>-statproc = icon_led_green.
          ELSE.
            <ls_object>-statproc = icon_led_yellow.
          ENDIF.
        ENDIF.

      ENDLOOP.  "Target language
    ENDLOOP.    "LXE Sub-Objects
  ENDLOOP.      "Objects

  MESSAGE i398 WITH text-ex1 space space space DISPLAY LIKE 'S'. "Executed with success
ENDFORM.                    " COPY_TRANSLATIONS
*---------------------------------------------- Form DOWNLOAD_TEMPLATE *
FORM download_template TABLES lt_objects TYPE STANDARD TABLE. "#EC PF_NO_TYPE
  DATA:
    ls_object     LIKE LINE OF gt_objects,    "Objects to transport
    ls_tlanguage  LIKE LINE OF gt_languages,  "Languages Informations
    ls_slanguage  LIKE LINE OF gt_languages,
    ls_tlang      LIKE LINE OF so_tlang,      "Target Languages
    lt_pcx_s1     TYPE TABLE OF lxe_pcx_s1,   "Text Pairs
    ls_objs_colob LIKE LINE OF gt_objs_colob, "LXE Objects
    ls_pcx_s1     LIKE LINE OF lt_pcx_s1.
  DATA:
    lv_filename          TYPE string,     "File
    lv_path              TYPE string,
    lv_fullpath          TYPE string,
    lv_window_title      TYPE string,     "Popup Windows Title
    lv_default_file_name TYPE string,     "Default file
    lv_user_action       TYPE i,          "User Actions
    lv_object            TYPE string,     "Excel object
    lv_lxe_object        TYPE string,     "LXE object
    lv_column            TYPE i VALUE 1,  "Excel Columns
    lv_row               TYPE i VALUE 1,  "Excel rows
    lv_row_lang          TYPE i,          "Excel languages rows
    lv_lang_txt          TYPE string,     "Language description
    lv_add_row           TYPE abap_bool,  "Add new row to excel flag
    lv_row_init          TYPE i,          "Language start row
    lv_obj_text          TYPE lxe0060lin. "Texts for Object Attributes
  DATA:
    lo_application TYPE ole2_object,  "OLE Automation Controller: OLE Typen
    lo_workbook    TYPE ole2_object,
    lo_worksheet   TYPE ole2_object,
    lo_column      TYPE ole2_object.

*---------- Save file dialog ----------*
  lv_window_title = text-g02. "Download Template
  READ TABLE gt_languages INTO ls_slanguage WITH KEY laiso = p_slang.
  CONCATENATE p_obj_n ls_slanguage-r3_lang text-f01 gv_tlangs INTO lv_default_file_name SEPARATED BY space.
  CONCATENATE lv_default_file_name '.' gc_excel_ext INTO lv_default_file_name.

  CALL METHOD cl_gui_frontend_services=>file_save_dialog
    EXPORTING
      window_title         = lv_window_title
      default_extension    = gc_excel_ext
      default_file_name    = lv_default_file_name
      prompt_on_overwrite  = abap_false
    CHANGING
      filename             = lv_filename
      path                 = lv_path
      fullpath             = lv_fullpath
      user_action          = lv_user_action
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      OTHERS               = 4.
  IF sy-subrc IS NOT INITIAL.
    MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'. "Executed with errors
    RETURN.
  ENDIF.

*---------- Create excel ----------*
  IF lv_user_action IS INITIAL.
    CREATE OBJECT lo_application 'excel.application'.       "#EC NOTEXT
    IF sy-subrc IS NOT INITIAL.
      MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'. "Executed with errors
      RETURN.
    ENDIF.
    SET PROPERTY OF lo_application 'visible' = 0.           "#EC NOTEXT
*---------- Create Workbook ----------*
    CALL METHOD OF lo_application 'Workbooks' = lo_workbook.  "Get Workbook
    SET PROPERTY OF lo_application 'SheetsInNewWorkbook' = 1.
    CALL METHOD OF lo_workbook 'Add'.                         "Create Workbook
*---------- Create Worksheet ----------*
    CALL METHOD OF lo_application 'Worksheets' = lo_worksheet "Get worksheet
      EXPORTING #1 = 1.
    CALL METHOD OF lo_worksheet 'Activate'.                   "Activate worksheet
    SET PROPERTY OF lo_worksheet 'Name' = text-f01.         "#EC NOTEXT
*---------- Create Header Column Object ----------*
    PERFORM create_cell USING lv_row lv_column gc_object abap_true CHANGING lo_worksheet.
*---------- Create Header Column LXE Object ----------*
    ADD 1 TO lv_column.
    PERFORM create_cell USING lv_row lv_column gc_lxe_object abap_true CHANGING lo_worksheet.
*---------- Create Header Column Text Pair ----------*
    ADD 1 TO lv_column.
    PERFORM create_cell USING lv_row lv_column gc_textkey abap_true CHANGING lo_worksheet.
*---------- Create Header Column LXE object description ----------*
    ADD 1 TO lv_column.
    PERFORM create_cell USING lv_row lv_column gc_desc abap_true CHANGING lo_worksheet.
*---------- Create Header Column Length ----------*
    ADD 1 TO lv_column.
    PERFORM create_cell USING lv_row lv_column gc_length abap_true CHANGING lo_worksheet.
*---------- Create Header Column Source Language column ----------*
    ADD 1 TO lv_column.
    CONCATENATE ls_slanguage-r3_lang ls_slanguage-text INTO lv_lang_txt SEPARATED BY space.
    PERFORM create_cell USING lv_row lv_column lv_lang_txt abap_true CHANGING lo_worksheet.

*---------- Create Header Columns Target Languages ----------*
    LOOP AT so_tlang INTO ls_tlang.
      READ TABLE gt_languages INTO ls_tlanguage WITH KEY laiso = ls_tlang-low. "Read language tecnical info
      IF sy-subrc IS INITIAL.
        ADD 1 TO lv_column.
        CONCATENATE ls_tlanguage-r3_lang ls_tlanguage-text INTO lv_lang_txt SEPARATED BY space.
        PERFORM create_cell USING lv_row lv_column lv_lang_txt abap_true CHANGING lo_worksheet.
      ENDIF.
    ENDLOOP.

*---------- Loop all selected objects ----------*
    LOOP AT lt_objects INTO ls_object.
      CONCATENATE ls_object-pgmid ls_object-object ls_object-obj_name INTO lv_object SEPARATED BY space.
      lv_add_row = abap_true. "Set add rows for new object
      lv_column = 6.          "Start text column
      lv_row_init = lv_row.   "Get init row

*---------- Create Column by target languague ----------*
      LOOP AT so_tlang INTO ls_tlang.                    "#EC CI_NESTED
        ADD 1 TO lv_column.
        CLEAR ls_tlanguage.
        READ TABLE gt_languages INTO ls_tlanguage WITH KEY laiso = ls_tlang-low. "Read language tecnical info

*---------- Create Rows by LXE object ----------*
        LOOP AT gt_objs_colob INTO ls_objs_colob WHERE pgmid    = ls_object-pgmid "#EC CI_NESTED
                                                   AND object   = ls_object-object
                                                   AND obj_name = ls_object-obj_name.
          CONCATENATE ls_objs_colob-objtype ls_objs_colob-objname INTO lv_lxe_object SEPARATED BY space.
*---------- Read texts ----------*
          REFRESH lt_pcx_s1.
          CALL FUNCTION 'LXE_OBJ_TEXT_PAIR_READ'
            EXPORTING
              t_lang    = ls_tlanguage-o_language
              s_lang    = ls_slanguage-o_language
              custmnr   = ls_objs_colob-custmnr
              objtype   = ls_objs_colob-objtype
              objname   = ls_objs_colob-objname
            TABLES
              lt_pcx_s1 = lt_pcx_s1.

*---------- Read objects description ----------*
          CLEAR lv_obj_text.
          CALL FUNCTION 'LXE_ATTOB_OBJECT_TYPE_TEXT_GET'    "#EC FB_RC
            EXPORTING
              obj_type      = ls_objs_colob-objtype
            IMPORTING
              obj_text      = lv_obj_text
            EXCEPTIONS
              no_text_found = 1
              OTHERS        = 2.

*---------- Create Rows (LXE Object) ----------*
          lv_row_lang = lv_row_init.
          LOOP AT lt_pcx_s1 INTO ls_pcx_s1.              "#EC CI_NESTED
            IF lv_add_row IS NOT INITIAL.
              ADD 1 TO lv_row.
              lv_row_lang = lv_row.
*---------- Create Rows (Object) ----------*
              PERFORM create_cell USING lv_row 1 lv_object abap_true CHANGING lo_worksheet.
*---------- Create Rows (LXE Object) ----------*
              PERFORM create_cell USING lv_row 2 lv_lxe_object abap_true CHANGING lo_worksheet.
*---------- Create Rows (Text Pair) ----------*
              PERFORM create_cell USING lv_row 3 ls_pcx_s1-textkey abap_true CHANGING lo_worksheet.
*---------- Create Rows (Description) ----------*
              PERFORM create_cell USING lv_row 4 lv_obj_text abap_true CHANGING lo_worksheet.
*---------- Create Rows (Length) ----------*
              PERFORM create_cell USING lv_row 5 ls_pcx_s1-unitmlt abap_true CHANGING lo_worksheet.
*---------- Create Rows (LXE Object Source Language) ----------*
              PERFORM create_cell USING lv_row 6 ls_pcx_s1-s_text abap_true CHANGING lo_worksheet.

            ELSE.
              ADD 1 TO lv_row_lang.
            ENDIF.
*---------- Create Rows (LXE Object Target languague ) ----------*
            PERFORM create_cell USING lv_row_lang lv_column ls_pcx_s1-t_text abap_false CHANGING lo_worksheet.
          ENDLOOP.  "Rows
        ENDLOOP.  "LXE object

        CLEAR lv_add_row.
      ENDLOOP.  "Column
    ENDLOOP.  "Selected objects

*---------- Format Columns ----------*
    CALL METHOD OF lo_application 'Columns' = lo_column.  "Get Column
    CALL METHOD OF lo_column 'Autofit'.                   "Set Column Autofit

    CALL METHOD OF lo_application 'COLUMNS' = lo_column "Get Column 2
      EXPORTING #1 = 2.
    SET PROPERTY OF lo_column 'ColumnWidth' = 1.

    CALL METHOD OF lo_application 'COLUMNS' = lo_column "Get Column 3
      EXPORTING #1 = 3.
    SET PROPERTY OF lo_column 'ColumnWidth' = 1.

*---------- Save excel worksheet ----------*
    CALL METHOD OF lo_worksheet 'SaveAs'                  "Save excel
      EXPORTING #1 = lv_fullpath.

    IF sy-subrc IS INITIAL.
*---------- Closes excel window ----------*
      CALL METHOD OF lo_workbook 'CLOSE'.                   "Close Workbook
      SET PROPERTY OF lo_application 'DisplayAlerts' = 0.
      CALL METHOD OF lo_application 'QUIT'.                 "End excel
      FREE OBJECT: lo_column, lo_worksheet, lo_workbook, lo_application.
      MESSAGE i398 WITH text-ex1 space space space DISPLAY LIKE 'S'. "Executed with success
    ELSE.
      MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'. "Executed with errors
      RETURN.
    ENDIF.
  ENDIF.

ENDFORM.                    " DOWNLOAD_TEMPLATE
*----------------------------------------------- Form UPLOAD_TEMPLATE *
FORM upload_template TABLES lt_objects TYPE STANDARD TABLE. "#EC PF_NO_TYPE
  DATA:
    ls_object          LIKE LINE OF gt_objects,                   "Object to transport
    lt_files           TYPE TABLE OF file_table,                  "Files names
    ls_file            LIKE LINE OF lt_files,
    lt_excel_data      TYPE TABLE OF alsmex_tabline,              "Rows for Table with Excel Data
    ls_excel_data      LIKE LINE OF lt_excel_data,
    lt_component_table TYPE cl_abap_structdescr=>component_table, "Component Description Table
    ls_component_table LIKE LINE OF lt_component_table,
    ls_tlanguage       LIKE LINE OF gt_languages,                 "Languages Informations
    ls_slanguage       LIKE LINE OF gt_languages,
    ls_objs_colob      LIKE LINE OF gt_objs_colob,                "LXE Objects
    lt_pcx_s1          TYPE TABLE OF lxe_pcx_s1,                  "Text Pairs
    ls_pcx_s1          LIKE LINE OF lt_pcx_s1,
    ls_tlang           LIKE LINE OF so_tlang.                     "Target Languages
  DATA:
    lo_data        TYPE REF TO data,                "Generic data
    lo_data_line   TYPE REF TO data,
    lo_excel_table TYPE REF TO cl_abap_tabledescr,  "Runtime Type Services
    lo_excel_type  TYPE REF TO cl_abap_structdescr.
  DATA:
    lv_title       TYPE string,      "Window title
    lv_rc          TYPE i,           "User action
    lv_filename    TYPE localfile,   "File Name
    lv_object      TYPE string,      "Object
    lv_pstatus     TYPE lxestatprc,  "Process Status
    lv_stattrn     TYPE lxestattrn,  "Translation Status of an Object
    lv_lxe_object  TYPE string,      "LXE object
    lv_tlang_exist TYPE abap_bool,  "Target Language flag
    lv_modify      TYPE abap_bool.
  FIELD-SYMBOLS:
    <lt_excel_table> TYPE STANDARD TABLE,      "Dynamic excel
    <ls_excel_table> TYPE any,
    <field>          TYPE any,                 "Field pointer
    <ls_pcx_s1>      LIKE LINE OF lt_pcx_s1,   "Text Pairs
    <ls_object>      LIKE LINE OF gt_objects.  "Objects to translate

*---------- Open file dialog ----------*
  lv_title = text-g01.  "Upload Template
  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = lv_title
      default_filename        = '*.xlsx'
    CHANGING
      file_table              = lt_files
      rc                      = lv_rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc IS NOT INITIAL.
    MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'. "Executed with errors
    RETURN.
  ENDIF.

*---------- Read excel ----------*
  IF lv_rc = 1.
    READ TABLE lt_files INTO ls_file INDEX 1.
    IF sy-subrc IS INITIAL.
*---------- Upload excel ----------*
      lv_filename = ls_file-filename.
      CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
        EXPORTING
          filename                = lv_filename
          i_begin_col             = 1
          i_begin_row             = 1
          i_end_col               = 10000
          i_end_row               = 10000
        TABLES
          intern                  = lt_excel_data
        EXCEPTIONS
          inconsistent_parameters = 1
          upload_ole              = 2
          OTHERS                  = 3.
      IF sy-subrc IS NOT INITIAL OR lt_excel_data IS INITIAL.
        MESSAGE i398 WITH text-m02 space space space DISPLAY LIKE 'E'. "Error opening file
        RETURN.
      ENDIF.
*---------- Create dynamic table structure ----------*
      LOOP AT lt_excel_data INTO ls_excel_data WHERE row = 1.
        CASE ls_excel_data-col.
          WHEN 1. "Object
            IF ls_excel_data-value <> gc_object.
              MESSAGE i398 WITH text-m11 space space space DISPLAY LIKE 'E'. "File not valid
              RETURN.
            ENDIF.
            ls_component_table-name = gc_object.  "Object Type
            ls_component_table-type = cl_abap_elemdescr=>get_c( '48' ).

          WHEN 2. "LXE Object
            IF ls_excel_data-value <> gc_lxe_object.
              MESSAGE i398 WITH text-m11 space space space DISPLAY LIKE 'E'. "File not valid
              RETURN.
            ENDIF.
            ls_component_table-name = gc_lxe_object. "LXE Object
            ls_component_table-type = cl_abap_elemdescr=>get_c( '75' ).

          WHEN 3. "Text Pair
            IF ls_excel_data-value <> gc_textkey.
              MESSAGE i398 WITH text-m11 space space space DISPLAY LIKE 'E'. "File not valid
              RETURN.
            ENDIF.
            ls_component_table-name = gc_textkey.  "Text pair
            ls_component_table-type = cl_abap_elemdescr=>get_c( '32' ).

          WHEN 4. "Description
            IF ls_excel_data-value <> gc_desc.
              MESSAGE i398 WITH text-m11 space space space DISPLAY LIKE 'E'. "File not valid
              RETURN.
            ENDIF.
            ls_component_table-name = gc_desc.  "Description
            ls_component_table-type = cl_abap_elemdescr=>get_c( '60' ).

          WHEN 5. "Length
            IF ls_excel_data-value <> gc_length.
              MESSAGE i398 WITH text-m11 space space space DISPLAY LIKE 'E'. "File not valid
              RETURN.
            ENDIF.
            ls_component_table-name = gc_length.  "Length
            ls_component_table-type = cl_abap_elemdescr=>get_c( '10' ).

          WHEN 6. "Source language
            READ TABLE gt_languages INTO ls_slanguage WITH KEY laiso = p_slang.
            IF ls_excel_data-value(2) <> ls_slanguage-r3_lang.
              MESSAGE i398 WITH text-m12 space space space DISPLAY LIKE 'E'. "File source language not valid
              RETURN.
            ENDIF.
            ls_component_table-name = ls_excel_data-value(2).  "Languagues
            ls_component_table-type = cl_abap_elemdescr=>get_c( '132' ).

          WHEN OTHERS. "Target languages
            READ TABLE gt_languages INTO ls_tlanguage WITH KEY r3_lang = ls_excel_data-value(2).
            IF sy-subrc IS INITIAL.
              lv_tlang_exist = abap_true.
            ENDIF.

            ls_component_table-name = ls_excel_data-value(2).  "Languagues
            ls_component_table-type = cl_abap_elemdescr=>get_c( '132' ).
        ENDCASE.

        APPEND ls_component_table TO lt_component_table.
        CLEAR ls_component_table.
      ENDLOOP.

      IF lv_tlang_exist IS INITIAL.
        MESSAGE i398 WITH text-m13 space space space DISPLAY LIKE 'E'. "File target languages not valid
        RETURN.
      ENDIF.

*---------- Create dynamic table ----------*
      IF lt_component_table IS NOT INITIAL.
        lo_excel_type = cl_abap_structdescr=>create( lt_component_table ).
        lo_excel_table = cl_abap_tabledescr=>create( p_line_type  = lo_excel_type
                                                     p_table_kind = cl_abap_tabledescr=>tablekind_std
                                                     p_unique     = abap_false ).
        CREATE DATA lo_data TYPE HANDLE lo_excel_table.
        ASSIGN lo_data->* TO <lt_excel_table>.
        IF <lt_excel_table> IS ASSIGNED.
          CREATE DATA lo_data_line LIKE LINE OF <lt_excel_table>.
          ASSIGN lo_data_line->* TO <ls_excel_table>.
        ENDIF.
      ENDIF.

      IF <lt_excel_table> IS NOT ASSIGNED OR <ls_excel_table> IS NOT ASSIGNED.
        MESSAGE i398 WITH text-ex2 space space space DISPLAY LIKE 'E'. "Executed with errors
        RETURN.
      ENDIF.
*---------- Fill dynamic table with excel data ----------*
      LOOP AT lt_excel_data INTO ls_excel_data WHERE row > 1.
        AT NEW row.                                     "#EC AT_LOOP_WH
          APPEND INITIAL LINE TO <lt_excel_table> ASSIGNING <ls_excel_table>.
        ENDAT.

        READ TABLE lt_component_table INTO ls_component_table INDEX ls_excel_data-col.
        IF sy-subrc IS INITIAL.
          ASSIGN COMPONENT ls_component_table-name OF STRUCTURE <ls_excel_table> TO <field>.
          IF <field> IS NOT ASSIGNED.
            CONTINUE.
          ENDIF.
        ENDIF.
        <field> = ls_excel_data-value.
      ENDLOOP.

      IF <lt_excel_table> IS INITIAL.
        MESSAGE i398 WITH text-m04 space space space DISPLAY LIKE 'W'. "No objects selected
      ENDIF.

*---------- Translate upload excel data ----------*
*---------- Loop all selected objects ----------*
      LOOP AT lt_objects INTO ls_object.                 "#EC CI_NESTED
        CONCATENATE ls_object-pgmid ls_object-object ls_object-obj_name INTO lv_object SEPARATED BY space.
        READ TABLE <lt_excel_table> ASSIGNING <ls_excel_table> WITH KEY (gc_object) = lv_object.  "Check if object exist in file
        IF sy-subrc IS NOT INITIAL. "Object not found in file
          CONTINUE.
        ENDIF.

*---------- Loop LXE Sub-Objects ----------*
        LOOP AT gt_objs_colob INTO ls_objs_colob WHERE pgmid    = ls_object-pgmid "#EC CI_NESTED
                                                   AND object   = ls_object-object
                                                   AND obj_name = ls_object-obj_name.

          CONCATENATE ls_objs_colob-objtype ls_objs_colob-objname INTO lv_lxe_object SEPARATED BY space.
          READ TABLE <lt_excel_table> ASSIGNING <ls_excel_table> WITH KEY (gc_object)     = lv_object "Check if object exist in file
                                                                          (gc_lxe_object) = lv_lxe_object.
          IF sy-subrc IS NOT INITIAL. "Object not found in file
            CONTINUE.
          ENDIF.

*---------- Loop selected target language ----------*
          LOOP AT so_tlang INTO ls_tlang.                "#EC CI_NESTED
            CLEAR ls_tlanguage.
            READ TABLE gt_languages INTO ls_tlanguage WITH KEY laiso = ls_tlang-low. "Read language tecnical info
            READ TABLE lt_component_table TRANSPORTING NO FIELDS WITH KEY name = ls_tlanguage-r3_lang. "Check if target language exist
            IF sy-subrc IS NOT INITIAL. "Target Language found
              CONTINUE.
            ENDIF.

*---------- Read texts ----------*
            CLEAR lv_pstatus.
            REFRESH lt_pcx_s1.
            CALL FUNCTION 'LXE_OBJ_TEXT_PAIR_READ'
              EXPORTING
                t_lang    = ls_tlanguage-o_language
                s_lang    = ls_slanguage-o_language
                custmnr   = ls_objs_colob-custmnr
                objtype   = ls_objs_colob-objtype
                objname   = ls_objs_colob-objname
                read_only = space
              IMPORTING
                pstatus   = lv_pstatus
              TABLES
                lt_pcx_s1 = lt_pcx_s1.
            IF lv_pstatus <> 'S' OR lt_pcx_s1 IS INITIAL. "Not Successful or empty
              CONTINUE.
            ENDIF.

*---------- Update and check Overwrite ----------*
            CLEAR lv_modify.
            LOOP AT lt_pcx_s1 ASSIGNING <ls_pcx_s1>.     "#EC CI_NESTED
              IF p_ow IS INITIAL AND <ls_pcx_s1>-t_text IS NOT INITIAL. "Check Overwrite
                CONTINUE.
              ENDIF.

              READ TABLE <lt_excel_table> ASSIGNING <ls_excel_table> WITH KEY (gc_object)     = lv_object
                                                                              (gc_lxe_object) = lv_lxe_object
                                                                              (gc_textkey)    = <ls_pcx_s1>-textkey.
              IF sy-subrc IS INITIAL.
                ASSIGN COMPONENT ls_tlanguage-r3_lang OF STRUCTURE <ls_excel_table> TO <field>.
                IF sy-subrc IS INITIAL.
                  IF <field> IS NOT INITIAL.
                    <ls_pcx_s1>-t_text = <field>.
                    lv_modify = abap_true.
                  ENDIF.
                ENDIF.
              ENDIF.
            ENDLOOP.

*---------- Update texts ----------*
            IF lv_modify IS NOT INITIAL.
              CALL FUNCTION 'LXE_OBJ_TEXT_PAIR_WRITE'
                EXPORTING
                  t_lang    = ls_tlanguage-o_language
                  s_lang    = ls_slanguage-o_language
                  custmnr   = ls_objs_colob-custmnr
                  objtype   = ls_objs_colob-objtype
                  objname   = ls_objs_colob-objname
                TABLES
                  lt_pcx_s1 = lt_pcx_s1.
*---------- Create proposal and check status ----------*
              LOOP AT lt_pcx_s1 INTO ls_pcx_s1 WHERE t_text IS NOT INITIAL. "#EC CI_NESTED
                CALL FUNCTION 'LXE_PP1_PROPOSAL_EDIT_SE63'
                  EXPORTING
                    t_lang         = ls_tlanguage-o_language
                    s_lang         = ls_slanguage-o_language
                    custmnr        = ls_objs_colob-custmnr
                    objtype        = ls_objs_colob-objtype
                    domatyp        = ls_objs_colob-domatyp
                    domanam        = ls_objs_colob-domanam
                    pcx_s1         = ls_pcx_s1
                    direct_command = 'ASTX'
                    direct_status  = '69'.
              ENDLOOP.  "Proposal
            ENDIF.

*---------- Check and update log translation status ----------*
            READ TABLE gt_objects ASSIGNING <ls_object> WITH KEY pgmid    = ls_object-pgmid
                                                                 object   = ls_object-object
                                                                 obj_name = ls_object-obj_name.
            IF sy-subrc IS INITIAL.
              CLEAR lv_stattrn.
              CALL FUNCTION 'LXE_OBJ_TRANSLATION_STATUS2'
                EXPORTING
                  t_lang  = ls_tlanguage-o_language
                  s_lang  = ls_slanguage-o_language
                  custmnr = ls_objs_colob-custmnr
                  objtype = ls_objs_colob-objtype
                  objname = ls_objs_colob-objname
                IMPORTING
                  stattrn = lv_stattrn.
*---------- Process Translation status ----------*
              IF lv_stattrn = 'T'.  "Translated
                IF <ls_object>-statproc <> icon_led_yellow.
                  <ls_object>-statproc = icon_led_green.
                ENDIF.
              ELSE.
                <ls_object>-statproc = icon_led_yellow.
              ENDIF.
            ENDIF.

          ENDLOOP.  "Target Languagens
        ENDLOOP.  "LXE Sub-Objects
      ENDLOOP.  "Objects

      MESSAGE i398 WITH text-ex1 space space space DISPLAY LIKE 'S'.  "Executed with success
    ENDIF.  "Open excel
  ENDIF.  "Read excel

ENDFORM.                    " UPLOAD_TEMPLATE
*--------------------------------------------------- Form CREATE_CELL *
FORM create_cell USING p_row_num TYPE i p_cell_num TYPE i p_value TYPE any p_bold TYPE abap_bool
              CHANGING c_sheet   TYPE ole2_object.
  DATA:
    lo_cell TYPE ole2_object,
    e_bold  TYPE ole2_object.

*---------- Create Excel cell ----------*
  CALL METHOD OF c_sheet 'Cells' = lo_cell  "Get Cell
    EXPORTING #1 = p_row_num #2 = p_cell_num.
  SET PROPERTY OF lo_cell 'Value' = p_value.                "#EC NOTEXT
  IF p_bold IS NOT INITIAL.
    GET PROPERTY OF lo_cell 'Font' = e_bold.                "#EC NOTEXT
    SET PROPERTY OF e_bold 'Bold' = 1.                      "#EC NOTEXT
  ENDIF.
ENDFORM.                    " CREATE_CELL

*---------------------------------- CLASS HANDLE EVENTS IMPLEMENTATION *
CLASS lcl_handle_events IMPLEMENTATION.
*-------------------------------------------------------- User command *
  METHOD on_user_command.
    CONSTANTS:
      lc_copy TYPE string VALUE 'COPY', "Copy original to targets
      lc_down TYPE string VALUE 'DOWN', "Download Template
      lc_up   TYPE string VALUE 'UP',   "Upload Template
      lc_tr   TYPE string VALUE 'TR'.   "Transport request

    CHECK e_salv_function = lc_copy OR e_salv_function = lc_down OR
          e_salv_function = lc_up   OR e_salv_function = lc_tr.

    DATA lr_selections TYPE REF TO cl_salv_selections. "ALV Selections
    DATA:
      lt_rows    TYPE salv_t_row,            "ALV Rows
      ls_row     TYPE i,
      lt_objects TYPE TABLE OF gty_objects.  "Objects to translate
    DATA lv_answer TYPE c.
    FIELD-SYMBOLS <ls_object> LIKE LINE OF gt_objects.  "Objects to translate

    TRY.
*---------- Get selected lines ----------*
        lr_selections = go_objects->get_selections( ).
        lt_rows = lr_selections->get_selected_rows( ).

*---------- Get selected objects ----------*
        LOOP AT lt_rows INTO ls_row.
          READ TABLE gt_objects ASSIGNING <ls_object> INDEX ls_row.
          IF sy-subrc IS INITIAL AND <ls_object>-status = icon_led_green. "Object valid if status OK
            CLEAR <ls_object>-statproc.                                   "Clear translation status for reprocessing
            APPEND <ls_object> TO lt_objects.                             "Add selected object for processing
          ENDIF.
        ENDLOOP.

        IF lt_objects IS NOT INITIAL. "Objects selected
*---------- Confirmation for copy or update ----------*
          IF e_salv_function = lc_copy OR e_salv_function = lc_up.
            CALL FUNCTION 'POPUP_TO_CONFIRM'   "#EC FB_RC "#EC CI_SUBRC
              EXPORTING
                titlebar       = text-f01
                text_question  = text-f02
              IMPORTING
                answer         = lv_answer
              EXCEPTIONS
                text_not_found = 1
                OTHERS         = 2.
            IF lv_answer <> '1'. RETURN. ENDIF.
          ENDIF.

*---------- Execution ----------*
          CASE e_salv_function.
*---------- Copy original to targets ----------*
            WHEN lc_copy.
              PERFORM copy_translations TABLES lt_objects.
*---------- Download Template ----------*
            WHEN lc_down.
              PERFORM download_template TABLES lt_objects.
*---------- Upload Template ----------*
            WHEN lc_up.
              PERFORM upload_template TABLES lt_objects.
*---------- Transport request ----------*
            WHEN lc_tr.
              PERFORM create_transport TABLES lt_objects.
          ENDCASE.

*---------- ALV Refresh for copy or update ----------*
          IF e_salv_function = lc_copy OR e_salv_function = lc_up.
            go_objects->refresh( ).
          ENDIF.

        ELSE.
          MESSAGE i398 WITH text-m04 space space space. "No objects selected
        ENDIF.

      CATCH cx_root INTO go_exp.                         "#EC CATCH_ALL
        gv_msg_text = go_exp->get_text( ).
        MESSAGE s398 WITH gv_msg_text space space space DISPLAY LIKE 'E'. "Critical error
    ENDTRY.
  ENDMETHOD.                    "on_user_command
*-------------------------------------------------------- Line dbclick *
  METHOD on_double_click.
    DATA:
      lt_spopli   TYPE TABLE OF spopli,       "Language infos
      ls_spopli   LIKE LINE OF lt_spopli,
      ls_language LIKE LINE OF gt_languages,
      ls_e071     TYPE e071,                  "Change & Transport System: Object Entries of Requests/Tasks
      ls_object   LIKE LINE OF gt_objects.    "Objects to transport
    DATA:
      lv_answer TYPE c,     "User Answer
      lv_tlang  TYPE spras. "Target Translation Language

*---------- Get selected line ----------*
    READ TABLE gt_objects INTO ls_object INDEX row.
    IF sy-subrc IS INITIAL AND ls_object-status = icon_led_green.
      TRY.
          IF lines( so_tlang ) = 1. "Only one target language
            LOOP AT gt_languages INTO ls_language WHERE laiso <> p_slang.
              lv_tlang = ls_language-laiso.
              EXIT.
            ENDLOOP.

          ELSE. "More that one target language
*---------- Fill popup radio buttons ----------*
            LOOP AT gt_languages INTO ls_language WHERE laiso <> p_slang.
              CONCATENATE ls_language-r3_lang ls_language-text INTO ls_spopli-varoption
                SEPARATED BY space.
              APPEND ls_spopli TO lt_spopli.
            ENDLOOP.
*---------- Please select target language to Edit ----------*
            CALL FUNCTION 'POPUP_TO_DECIDE_LIST'
              EXPORTING
                textline1          = text-d01
                titel              = text-d02
              IMPORTING
                answer             = lv_answer
              TABLES
                t_spopli           = lt_spopli
              EXCEPTIONS
                not_enough_answers = 1
                too_much_answers   = 2
                too_much_marks     = 3
                OTHERS             = 4.
            IF sy-subrc IS INITIAL AND lv_answer <> 'A'.  "Selected
              READ TABLE lt_spopli INTO ls_spopli WITH KEY selflag = abap_true.                   "Get selected radio button
              READ TABLE gt_languages INTO ls_language WITH KEY r3_lang = ls_spopli-varoption(2). "Get selected language
              lv_tlang = ls_language-laiso.
            ENDIF.
          ENDIF.

*---------- Edit Target Translation Language ----------*
          IF lv_tlang IS NOT INITIAL.
            MOVE-CORRESPONDING ls_object TO ls_e071.
            CALL FUNCTION 'LXE_OBJ_CALL_WL_SE63'
              EXPORTING
                s_lang  = p_slang
                t_lang  = lv_tlang
                do_eval = abap_true
                e071    = ls_e071.
          ENDIF.

        CATCH cx_root INTO go_exp.                       "#EC CATCH_ALL
          gv_msg_text = go_exp->get_text( ).
          MESSAGE s398 WITH gv_msg_text space space space DISPLAY LIKE 'E'. "Critical error
      ENDTRY.
    ENDIF.
  ENDMETHOD.                    "on_double_click
ENDCLASS.                    "lcl_handle_events IMPLEMENTATION
