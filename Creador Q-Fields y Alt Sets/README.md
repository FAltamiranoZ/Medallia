Instrucciones:

0.- Descargar el programa y colocarlo en un directorio sin espacios; es decir, que las carpetas que lo contegan no tengan espacios en sus nombre.
Por ejemplo: “C:\Descargas” es un directorio válido “C:\Carpeta1\Carpeta2\Carpeta3” también es válido, pero “C:\Carpeta1\Carpeta2\Carpeta 3” no es válido por el espacio en el nombre de “Carpeta 3” 

1.- Ajustar el archivo de excel adjunto para que esté acorde a la encuesta.
Se puede cambiar el nombre del archivo siempre y cuando no se agreguen espacios extras.
No se pueden agregar columnas extras al archivo de Excel (si es necesario, se pueden agregar hasta la derecha, pero no prometo que no explote).
El nombre de la compañía debe de ser el de la instancia (es decir, el que se usa en la url y no tiene espacios).
Se pueden agregar tantas filas como se quieran mientras se mantenga el formato de las páginas y se enumeren todas las preguntas (no se deben enumerar aquellas que son texto).
Para crear q fields, siempre deben estar llenas las columnas de Nombre Pregunta, Pregunta o texto, posibles respuestas y terminación; las columnas de Nombre en export, ¿Pregunta otro?, Altset, ace y duplicate checking son opcionales.
En caso de querer utilizar las columnas de ¿Pregunta otro?, ACE y Duplicate Checking, estas deben de rellenarse con "true", de lo contrario, se dejan en blanco.
En caso de no rellenar la celda de altset, se creará un altset nuevo con base en la columna de Posibles Respuestas. Se creará un altdb por cada línea en la celda. Si se le quieren poner labels a los altsets, es necesario hacerlo utilizando el carácter “|” y dejando los demás normal. Por ejemplo:
•	1|nada satisfecho
•	2
•	3
•	4
•	5|Muy satisfecho
Nota, si se desea compartir con el cliente, se pueden quitar desde las columnas “H” hasta la “L” y compartirlo como spec normal.

2.- Ejecutar el programa (tarda de 20 a 40 segundos en abrir)
Ingresar la dirección del archivo, este debe utilizar el formato .xls. Por ejemplo: “C:\nombreCarpeta\SpecEncuesta.xls”
Ingresar el nombre de la hoja de excel al que haga referencia (por ejemplo, "Hoja 1")
Hacer clic en generar

3.- Abrir los Bulks generados en la misma carpeta en la que se encuentra el programa
Introducir el archivo de bulk de los altsets (en caso de que se haya creado) en Medallia y procesarlo
En caso de que se hayan creado nuevos altsets, actualizar el archivo de bulk de los qfields con los ids en la columna P, las celdas que deben actualizarse tienen la leyenda “Poner Aquí el altset generado al procesar el archivo de altset spec”.
Introducir el bulk de los qfields en Medallia y procesarlo (si se agregaron preguntas de "Otro" no introducir el campo que tiene un other question si no hasta después en un segundo bulk)

4.- Ser feliz como lombriz
