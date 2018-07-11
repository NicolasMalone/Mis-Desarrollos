# VolcadoDeAdjuntos

1.¿Que realiza el programa?
Primero se accedera a outlook (Aplicacion Windows) y se buscara un buzon concreto en donde se encontraran los correos que llegan cada dia, una vez situados en el buzon y/o carpeta deseada, se procede a aplicar un filtro por fecha mas nueva a todos los correos de la carpeta, ya ordenados se procede a buscar los correos los cuales tengan como asunto un nombre en particular (Estos son 4), una vez ubicados se comprubea que tengan un adjunto, si este tiene alguno se procede a descargarlo en la ruta "D:\vdata\".

Tras tener todos los adjuntos estos son abiertos uno en uno y leido, pasando sus datos a un array (Como habia que cambiar algunos 	  formatos de datos se opto por usar un array), tras esto los datos del array son volcados a la base de datos.

Una vez los datos esten en la base de datos se actualizara el campo tiempo de las tablas "cotc" y "cota", el tiempo se obtendra buscando esta secuencia en el campo comentario --> [un tiempo], es decir el programa buscara un "[" y a partir de ahi verificara si hay numeros los cuales seran los tiempos, una vez que se encuentre con esto "]" el valor numero que ha encontrado lo pondra en el campo tiempo correspondiente a ese id de registro.
