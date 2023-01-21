

function menu() {
    /* --------------------------------------------------------------------------
    ' Procedure : menu
    ' Purpose   : Creates menus
    ' --------------------------------------------------------------------------*/

    var ui = SpreadsheetApp.getUi(); 
    ui.createMenu('Speeder') //crea menú
        .addItem('option1', 'procedure1')
        .addItem('option2', 'procedure2')
        .addToUi();
  }