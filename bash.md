# Bash - Linux system command line

In this document you will find all the documantation associated with Bash and linux system command line to make simple and frequent operations.

## Common procedures



## File management

``cd``: ir a la carpeta de inicio

``pwd``: ruta actual directorio donde estas

``ls``: ver contenido de un directorio

``cat <file>``: listar el contenido de un archivo en sdout

``cat > <file>``: crea nuevo archivo

``cat <file1> <file2> > <file3>``: une dos archivos y da salida en nuevoarchivo3

``cp <file> <path>``: copia archivo en dir actual a nuevo dir

``mv <file> <path>``: mueve archivo

``mkdir <dir>/<dir2>/<dir3>``: crea nuevo directorio

``rmdir``: eliminar dir vacios

``rm``: elimina directorio y contenido

``rm-r``: elimina solo el dir y deja archivos intactos

``touch <dir>/<file>``: crea archivo en blanco

``locate -i <file>``: localiza un archivo

``find /home/ -name archivo.txt``: busca archivo dentro de un directorio y sus subdir

``grep <string> <file>``: busca palabra dentro de un archivo

``diff <file1> <file2>``: compara linea por linea dos archivos y devuelve diferencias

``chmod``:cambiar los permisos sobre archivos y directorios

``chow``: <newUserName> <file>: cambiar propiedad de un archivo de un usario a otro

``echo <string> >> <file>``: mover datos hacia un archivo

``zip <file>``: zipear un archivo

``unzip <file>``: unzippear un archivo

## System management

``sudo``: comando para ejecutar los comandos en modo root. Ej: sudo + command

``sudo visudo``: ver el archivo/diretorio que controla sudo

``lsof <dir>``: ver que proceso está usando determinado archivo

``sudo -i``: ir hacia el directorio raiz

``sudo apt-get update``: actualizar los paquetes instalados

``sudo apt-get <package>``: instalar un paquete

``uname``: informacion del sistema y del kernel

``lsb_release -a``: version de ubuntu instalada

``jobs``: muestra trabajos actuales

``kill`` <flag> <PID>: matar proceso

``ps ux``: conocer PID

``ping <serverIP>``: verificar conexion a un servidor y tiempo respuesta

``wget <link>``: descargar archivo internet

``top``: monitorizar procesos y uso recursos

``history``: revisar historial de comandos utilizados

``man <commandName>``: instrucciones para un comando

``hostname -I``: conocer nombre de tu host/red

``useradd <userName>`` : agrega un nuevo usuario

``userdel <userName>``: eliminar un usuario

``<command1>;<command2>;<command3>``: ejecutar varios comandos a la vez sin importar secuencia

``<command1>&&<command2>&&<command3>``: idem pero sólo si el primer comando es exitorso

``whereis <applicationName>``: te da la ubicación de una aplicacion

``service --status-all``: listar servicios

## Console usage

CTLR + ALT + T: Open the console

Ctrl + C: detener comando en ejecucion de forma segura

Ctrl + Z: forzar detencion

TAB: Autocompletar

Ctlr + S: congelar terminal

Ctlr + Q: descongelar terminal

Ctlr + A: comienzo de la linea

Ctlr + E: final de la linea

SHIFT + AVPAG / REPAG subir y bajar para leer la terminal

``clear``: limpiar terminal
