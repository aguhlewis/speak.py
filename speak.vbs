Dim Message, speech
SpeaksMessage=InputBox("Enter text","Speak")
Set speech=CreateObject("sapi.spvoice")
Speech.Speak SpeaksMessage