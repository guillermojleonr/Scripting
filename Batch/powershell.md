# PowerShell personal documentation

- Abrir archivo

``Invoke-Item “rutaarchivo”``

- Ejecutar apps UWP

``Start-Process “app:”``
 
- Obtener ayuda

Get-Help
Get-Help <comando>
Get-Help <comando> -Full
Get-Help <comando> -Example
Get-Help *

- Actualizar librerías de ayuda

Update-Help

- Buscar comandos

Get-Command -Name Get-Command
Get-Command -CommandType <tipo>

- Muestra info sobre un directorio

Get-Item “ruta”

- Limpiar consola

Clear

- Ver contenido archivo

Get-Content “ruta archivo con ext”

- Copiar y Pegar un archivo

Copy-Item “ruta-archivo” -Destination “rutadestino”

- Borrar un archivo

Remove-Item “ruta”

- Gestionar servicios

Get-Service
Start-Service NombreServicio
Stop-Service NombreServicio
Suspend-Service NombreServicio
Resume-Service NombreServicio
Restart-Service NombreServicio

- Gestionar procesos

Get-Process NombreProceso
Start-Process NombreProceso
Stop-Process NombreProceso
Wait-Service NombreProceso

- Cambiar políticas de ejecución para scripts (permitir o no su ejecución)

Set-ExecutionPolicy Unrestricted
Set-ExecutionPolicy All Signed
Set-ExecutionPolicy Remote Signed
Set-ExecutionPolicy Restricted

- Información de un usuario

Get-LocalUser -name|fl

- Información del equipo

Get-WMIObject Win32_ComputerSystem

- Copiar todos los archivos del dir

ROBOCOPY “rutaorigen” “rutadestino”

- Copiar de forma recursiva carpetas con subdirectorios aunque estén vacíos.

ROBOCOPY ORIGEN DESTINO /E

- Copia de forma recursiva carpetas con subdirectorios pero no los vacios

ROBOCOPY ORIGEN DESTINO /S

- MIR modo espejo, Copia de forma recursiva pero al terminar se eliminan los archivos en el destino que ya no existen en el origen.

ROBOCOPY ORIGEN DESTINO /MIR


- Es posible indicar archivos específicos para ser copiados usando asteriscos de la siguiente forma:

ROBOCOPY ORIGEN DESTINO *.doc /E


- Opciones que permite el comando ROBOCOPY

/R:n	Numero de reintentos en caso de algún error.

/W:n	Tiempo de espera entre reintentos.

/MT:n	Realiza copias multiproceso, n especifica el número de hilos, el valor predeterminado es 8, n debe estar comprendido entre 1 y 128.

/MOV	Mueve archivos y los elimina del origen después de ser copiados.

/MOVE	Mueve archivos y carpetas y los elimina del origen después de ser copiados.

/V	Mostrar información detallada durante la copia.

/L	Hace una simulación, solo mostrar no copia.

/FP	Incluir ruta de acceso completa de los archivos en el resultado.

/NJH	No muestra el encabezado en la consola.

/NJS	No muestra el resumen final.

/Z	Copia archivos en modo reiniciable. Escribirá un registro en el archivo incompleto en caso de que la operación se vea interrumpida, para que en otra ejecución de Robocopy pueda continuarse por donde se dejó.

/MAX:n	Tamaño máximo de archivo, no se copian archivos mayores que el valor de n expresado en bytes.

/MIN:n	Tamaño mínimo de archivo, no se copian archivos menores que el valor de n expresado en bytes.

/MAXAGE:n	Antigüedad máxima de archivo, no se copian archivos mayores que el valor de n en días, puede usarse también fecha.

/MINAGE:n	Antigüedad mínima de archivo no se copian archivos menores que el valor de n en días, puede usarse también fecha.

/RH:hhmm-hhmm	Horas de ejecución, intervalo de horas en formato de 24 horas en que se debe iniciar la copia.

/LOG:log.txt	Permite guardar un informa con los datos de la copia efectuada en un archivo de texto.