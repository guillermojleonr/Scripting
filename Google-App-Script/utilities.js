/* --------------------------------------------------------------------------
  Author    : Guillermo Leon
' Website   : https://savingl.cl
  Purpose   : Store GAS some sample code
'--------------------------------------------------------------------------*/


function new_sheet1() {

    /* --------------------------------------------------------------------------
    ' Procedure : new_sheet
    ' Purpose   : Once you make a copy of sheet, run this function to clear the data. Notice that this function is related to a specific sheet
    '--------------------------------------------------------------------------*/

    const ss = SpreadsheetApp.getActive();
    const lastrow = ss.getDataRange().getValues().length -1;
    const ui = SpreadsheetApp.getUi()
    const inputpromt = ui.prompt('Introduce el número de inicio')
    let firstday = parseInt(inputpromt.getResponseText()) 
    const headerRange = ss.getRange('C4:CD4')
    let celltoset = ss.getRange('C4')
  
    //clearing headers
    ss.getRange('C4:CD4').clear({contentsOnly: true, skipFilteredRows: true}).setBackground(null);
    
    //clearing data
    ss.getRangeList([
        'C6:D'+ lastrow, 
        'F6:F'+ lastrow,
        'H6:I'+ lastrow,
        'K6:K'+ lastrow,
        'M6:N'+ lastrow,
        'P6:P'+ lastrow,
        'R6:S'+ lastrow,
        'U6:U'+ lastrow,
        'W6:X'+ lastrow,
        'Z6:Z'+ lastrow,
        'AB6:AC'+ lastrow,
        'AE6:AE'+ lastrow,
        'AG6:AH'+ lastrow,
        'AJ6:AJ'+ lastrow,
        'AL6:AM'+ lastrow,
        'AO6:AO'+ lastrow,
        'AQ6:AR'+ lastrow,
        'AT6:AT'+ lastrow,
        'AV6:AW'+ lastrow,
        'AY6:AY'+ lastrow,
        'BA6:BB'+ lastrow,
        'BD6:BD'+ lastrow,
        'BF6:BG'+ lastrow,
        'BI6:BI'+ lastrow,
        'BK6:BL'+ lastrow,
        'BN6:BN'+ lastrow,
        'BP6:BQ'+ lastrow,
        'BS6:BS'+ lastrow,
        'BU6:BV'+ lastrow,
        'BX6:BX'+ lastrow,
        'BZ6:CA'+ lastrow,
        'CC6:CC'+ lastrow,
        'CE6:CF'+ lastrow,
        'CH6:CH'+ lastrow,
        ]).clear({contentsOnly: true, skipFilteredRows: true})
  
    // setting headers
    for (i=1;i < 17; i++){
      celltoset.setValue(firstday)
      celltoset = celltoset.offset(0,5)
      firstday += 1
    }
  }


function new_records() {

    /* --------------------------------------------------------------------------
    ' Procedure : new_records()
    ' Purpose   : Once you make a copy of sheet, run this function to clear the data. Notice that this function is related to a specific sheet
    '--------------------------------------------------------------------------*/

    var sheet = SpreadsheetApp.getActiveSheet(); //active sheet object
  
    //Getting the last cell with values.
    var dvals = sheet.getRange("F1:F").getValues(); 
    var dlast = dvals.filter(String).length; 
  
    //Getting the ranges that we are going to copy
    var r_source = sheet.getRange(dlast,5,1,4); 
    var r_npaq = sheet.getRange(dlast,11); 
    var n_paq = sheet.getRange(dlast,11).getValue(); 
    var r_nrec = sheet.getRange(dlast,12); 
    
    // UI to prompt for the number of records that you need to create.
    var ui = SpreadsheetApp.getUi(); 
    var result = ui.prompt("Enter the record amount");
    var npaqs = result.getResponseText(); // gets the total shipments quantity entered
  
    //Performs copy-paste operation
    for (i=1; i < npaqs; i++) {
        dlast = dlast + 1
        n_paq = n_paq + 1
        r_npaq = sheet.getRange(dlast,11); 
        r_source.copyTo(sheet.getRange(dlast,5,1,4))
        r_nrec.copyTo(sheet.getRange(dlast,12))
        r_npaq.setValue(n_paq)
   }
  }
  
  
  function new_sheet2() {

    /* --------------------------------------------------------------------------
    ' Procedure : new_sheet2
    ' Purpose   : Restart the shipments worksheet.
    '--------------------------------------------------------------------------*/

    var ss = SpreadsheetApp.getActiveSpreadsheet(); // Activesheet Spreadsheet object, needed to create de trigger
    
    // Create a trigger for the new sheet.
    ScriptApp.newTrigger('onChange')
        .forSpreadsheet(ss)
        .onChange()
        .create();
  
    // Clear all the data except headers
    var dvals = ss.getRange("A1:A").getValues(); 
    var dlast = dvals.filter(String).length; 
    ss.getRange('A2:AE' + dlast).clear({contentsOnly: true, skipFilteredRows: true});
  
    
    var ssname = ss.getName(); 
    var ss = SpreadsheetApp.getActiveSheet(); //ActiveSheet object (Worksheet) is needed to make further operations, getActiveSpreadsheet() and getActiveSheet() aren't the same objects and do not have the same methods.
  
    // Set formulas
    ss.getRange(2,9).setFormula('Formula1') 
    ss.getRange(2,10).setFormula('Formula2') 
    ss.getRange(2,21,1,2).setValue(ssname) 
    
    // Copy-Paste the formulas on all rows
    ss.getRange('2:2').copyTo(ss.getRange('A2:AE1500'));
  
    // Create record ID, 
    var destination = ss.getRange('A2:A1500')
    var rango_correlativo = ss.getRange("A2:A3")
  
    var ui = SpreadsheetApp.getUi();
    var inputpromt = ui.prompt('Introduce the last ID from the previous spread sheet')
    var result1 = parseInt(inputpromt.getResponseText()) 
    var result2 = result1 + 1
   
    rango_correlativo.setValues([[result1], [result2]]) //first [] represents the macro array, second [] represents row arrays. 
    ss.getRange('A2:A3').autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  
  
  function colorSelectedRange() { 
    /* --------------------------------------------------------------------------
    ' Purpose   : Sets background to selected cell
    '--------------------------------------------------------------------------*/
    SpreadsheetApp.getActiveSheet().getActiveRange().setBackground('#d95b34');
  }
  
  
  
  function write_text_last_record() {

    /* --------------------------------------------------------------------------
    ' Purpose   : Writes a text on a specific field at the last record.
    ' Comments  : Does not work if you are not entering the last row.
    '--------------------------------------------------------------------------*/

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() // spreadsheet object
    let dvals = sheet.getRange("C1:C").getValues(); //column c values
    let dlast = dvals.filter(String).length;
    let r_observacion = sheet.getRange(dlast,15,1,1)
    
    r_observacion.setValue('text')
  }
  
  
  function warning_insert_rows(e){

    /*--------------------------------------------------------------------------
    ' Purpose   : Raise a warning to the user indicating that row inserts are not allowed.
    ' Comments  : This is just a warning, it doesn't prevent the user to be inserting new rows, a real solution would be to solve this by user privileges.'
    '--------------------------------------------------------------------------*/

    if(e.changeType == 'INSERT_ROW' || e.changeType == 'INSERT_COLUMN'){
      Browser.msgBox('No se permite insertar filas o columnas, presione CTRL + Z para deshacer y contacte al administrador')
    }
    else {
      return;
    }
  }
  

  
  function logissues() {
    /* --------------------------------------------------------------------------
    ' Procedure : logissues
    ' Purpose   : Verifies that a specific column formulas have correct references, specifically the match between field1 and field2 references.
    '--------------------------------------------------------------------------*/

   var ss = SpreadsheetApp.getActive();
   var lastrow = ss.getLastRow()
   var range1 = ss.getRange("A2:X" + lastrow)
   var values = range1.getValues()
   var issuesdetected = []
   var ui = SpreadsheetApp.getUi()
  
      for (i=0; i < values.length; i++) {
      let strdireccion = values[i][3]
      let splitString = strdireccion.split(' ')
      let comuna = splitString[splitString.length - 1]
      
       if (comuna !== values[i][23]) {
        issuesdetected.push(values[i][0]) 
      } 
    }
    ui.alert(issuesdetected.join())
  }