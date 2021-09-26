Dim message, sapi
message=InputBox("What do you want Kush to say?","Speak to Kush")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message