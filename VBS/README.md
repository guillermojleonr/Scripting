# Visual Basic Script
Some scripts in Visual Basic Script to do a few repetitive file-related tasks

Errores

- Type Mismatch 13: Puede ocurrir  mientras la carpeta de operación esté vacía, no tenga archivos. 

- No ejecuta la macro: La macro está contenida en la librería Lib.xlsm, el libro de macros personal debe contener dicha librería al momento de abrir el archivo de excel. Entre ellos tenemos los siguientes casos:

 1. No abre el libro de macros personal al abrir un archivo de excel: Puede que el libro de macros personal esté bloqueado por excel, para resolver este problema se debe navegar hasta: opciones de excel > complementos > elementos deshabilitados > ir. Habilitar el libro de macros nuevamente y verificar que se abre cada vez que se abre un archivo de excel.

 2. El libro de macros personal no retiene la librería. Esto se debe que al referenciar una librería en el libro de macros personal, para que ésta librería persista y no se pierda al cerrar el libro de macros personal, debe existir al menos un módulo en el libro de macros personal, no importa que no se use la librería en dicho módulo.