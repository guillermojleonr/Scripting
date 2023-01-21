/* --------------------------------------------------------------------------
  Author    : Guillermo Leon
' Website   : https://savingl.cl
  Purpose   : Manage procedures to be executed on specifid events
'--------------------------------------------------------------------------*/



function onOpen(e) {
    /* --------------------------------------------------------------------------
    ' Purpose   : Procedures to be executed on the open event
    '--------------------------------------------------------------------------*/
    menu() //create menu;
  }
  
  function onChange(e){
    /* --------------------------------------------------------------------------
    ' Purpose   : Procedures to be executed on the change event
    '--------------------------------------------------------------------------*/
    warning_insert_rows(e);
  }