<HTML>
<HEAD>
<!--#include file="clsUpload.asp"-->
</HEAD>
<BODY>
<FORM ACTION = "clsUploadTEST.asp" ENCTYPE="multipart/form-data" METHOD="POST">
Demo Input: <INPUT NAME = "Demo"></INPUT><P>
File Name: <INPUT TYPE=FILE NAME="txtFile"><P>
<INPUT TYPE = "SUBMIT" NAME="cmdSubmit" VALUE="SUBMIT">
</FORM><P>
<%

set o = new clsUpload
if o.Exists("cmdSubmit") then

'get client file name without path
sFileSplit = split(o.FileNameOf("txtFile"), "\")
sFile = sFileSplit(Ubound(sFileSplit))

o.FileInputName = "txtFile"
o.FileFullPath = Server.MapPath(".") & "\" & sFile
o.save

 if o.Error = "" then
	response.write "Success. File saved to  " & o.FileFullPath & ". Demo Input = " & o.ValueOf("Demo")
 else
	response.write "Failed due to the following error: " & o.Error
 end if

end if
set o = nothing
%>
</BODY>
</HTML>