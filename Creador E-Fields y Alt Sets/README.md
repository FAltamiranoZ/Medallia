Instrucciones:

0.- Descargar el programa y colocarlo en un directorio sin espacios; es decir, que las carpetas que lo contegan no tengan espacios en sus nombre.
Por ejemplo: “C:\Descargas” es un directorio válido “C:\Carpeta1\Carpeta2\Carpeta3” también es válido, pero “C:\Carpeta1\Carpeta2\Carpeta 3” no es válido por el espacio en el nombre de “Carpeta 3” 

1.- Ajustar el archivo de excel adjunto para que esté acorde al autoimporter.
Se puede cambiar el nombre del archivo siempre y cuando no se agreguen espacios extras.
No se pueden agregar columnas extras al archivo de Excel (si es necesario, se pueden agregar hasta la derecha, pero no prometo que no explote).
El nombre de la compañía debe de ser el de la instancia (es decir, el que se usa en la url, no tiene espacios y está en minúsculas).
Se pueden agregar tantas filas como se quieran mientras se mantenga el formato de las columnas
Solamente se crearán E-Fields cuando la celda de ¿Campo existente? Se encuentre vacía. Para su creación es necesario que estén rellenas las celdas de: Nombre de Columna recibida, Nombre en reporteo, Requerido y Tipo de Dato, las demás son opcionales.
Solamente se crearán Alt-Sets cuando la celda de ¿Campo existente? Se encuentre vacía, la celda de Tipo de Dato tenga el valor “Enumerado” y la celda de Altset contenga las opciones disponibles enlistadas en una línea nueva cada una. Para su creación es necesario que estén rellenas las celdas de: Nombre de Columna recibida, Nombre en reporteo, Requerido y Tipo de Dato y Altset, las demás son opcionales.
En caso de querer utilizar las columnas de ACE y Duplicate Checking, estas deben de rellenarse con S o N, en caso de dejarse en blanco, se asumirá que el valor es N.

2.- Ejecutar el programa (puede tardar hasta 30 segundos en abrir)
Ingresar la dirección del archivo, este debe utilizar el formato .xls. Por ejemplo: “C:\nombreCarpeta\SpecAutoImporter.xls”
Ingresar el nombre de la hoja de excel al que haga referencia (por ejemplo, "Hoja 1")
Hacer clic en generar

3.- Abrir los Bulks generados en la misma carpeta en la que se encuentra el programa
Introducir el archivo de bulk de los altsets (en caso de que se haya creado) en Medallia y procesarlo
En caso de que se hayan creado nuevos altsets, actualizar el archivo de bulk de los E-Fields con los ids de los nuevos AltSets en la columna K, las celdas que deben actualizarse tienen la leyenda “Poner Aquí el altset generado al procesar el archivo de altset spec”.
Introducir el bulk de los E-Fields en Medallia y procesarlo.

4.- Ser feliz como lombriz
