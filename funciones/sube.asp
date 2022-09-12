<!--#include file="xelupload.asp"-->
<%
Dim objUpload, objFich, strNombreFichero
Dim strNombre, strEdad

'Creamos el objeto 
set objUpload = new xelUpload

'Recibimos el formulario
objUpload.Upload()

'Mostramos total de ficheros recibidos
Response.Write ( objUpload.Ficheros.Count & " ficheros recibidos.")

'Y ahora mostramos los datos del fichero enviado:
'Lo sacamos a una variable por comodidad
set objFich = objUpload.Ficheros("fichero")

Response.Write ("<p>" & objFich.Nombre & "<br>")
Response.Write("Tamaño: " & objFich.Tamano & "<br>")
Response.Write("Tipo de contenido: " & objFich.TipoContenido & "</p>")

objFich.GuardarComo "archivo2.pdf",Server.MapPath("../ordenes")

set oFich = nothing
set objUpload = nothing

'Guardamos el fichero, con su nombre, en el directorio
'en el que se encuentra esta página




%>
