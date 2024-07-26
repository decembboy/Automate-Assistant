dim y
Dim x
do until x="stop"



x=inputbox("How can i help you??","BOT")




if x="google" then 

CreateObject("Wscript.shell").run "chrome.exe"



elseif x="d" then


CreateObject("Wscript.shell").run"D:\"


elseif x="e" then 

CreateObject("Wscript.shell").run"E:\"


elseif x="my computer" then 

CreateObject("Wscript.shell").run"C:\Users\Admin\Desktop"


elseif x="c" then 

CreateObject("Wscript.shell").run"C:\"


elseif x="whatsapp" then 

CreateObject("Wscript.shell").run"https://www.whatsapp.com/"

elseif x="python" then 

CreateObject("Wscript.shell").run"https://www.onlinegdb.com/online_python_compiler"


elseif x="ghy univ" then 

CreateObject("Wscript.shell").run"https://www.gauhati.ac.in/"


elseif x="ignou" then 

CreateObject("Wscript.shell").run"http://www.ignou.ac.in/"


elseif x="ms word" then 

CreateObject("Wscript.shell").run"""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Word 2007.lnk"""


elseif x="ms ex" then 

CreateObject("Wscript.shell").run"""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Excel 2007.lnk"""


elseif x="note" then 

CreateObject("Wscript.shell").run"notepad.exe"




elseif x="adobe" then 

CreateObject("Wscript.shell").run"""C:\Program Files\Adobe\Reader 9.0\Reader\AcroRd32.exe"""


elseif x="pic" then 

CreateObject("Wscript.shell").run"""C:\Users\Admin\Documents\pics\pics\DSC00017.jpg"""




elseif x="snipping" then 

createobject("wscript.shell").run"""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Snipping Tool.lnk""" 




elseif x="calc" then 

CreateObject("Wscript.shell").run"""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\calculator.lnk"""


elseif x="run" then 

CreateObject("Wscript.shell").run"""C:\Users\Admin\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Accessories\Run.lnk"""



elseif x="vehicle" then 

CreateObject("Wscript.shell").run"""https://parivahan.gov.in/parivahan/"""



elseif x="control" then 

CreateObject("Wscript.shell").run"""C:\Users\Admin\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Accessories\System Tools\Control Panel.lnk"""



elseif x="mail" then 

CreateObject("Wscript.shell").run"""https://mail.google.com/mail/u/0/#inbox"""







elseif x="how are you" then

createobject("sapi.spvoice").speak"i'm good thank u" 








elseif x="who build you" then

createobject("sapi.spvoice").speak"At first i was just an idea, then deborshi focused and put a lot of energy into the Bot application project. " 




elseif x="who is your boss" then

createobject("sapi.spvoice").speak"you're my boss. you give me purpose"






elseif x="stop" then
msgbox"the application has been closed",vbinformation,"Interpreter"



elseif x="who are you" then

createobject("sapi.spvoice").speak"i'm your Assistant, how can i help you"



elseif x="what is my name" then

createobject("sapi.spvoice").speak"what should i call you"
y=inputbox("Enter your name? ")
createobject("sapi.spvoice").speak"okay,i would call you"
createobject("sapi.spvoice").speak y


elseif x="do you have gf" then

createobject("sapi.spvoice").speak"This is one of those things we'd both have to agree on. i'd prefer to keep our relationship friendly . you know, romance makes me incredibly awkward."



elseif x="do you have bf" then


createobject("sapi.spvoice").speak "that's a secret, i can't tell you now"


elseif x="how were you made" then

createobject("sapi.spvoice").speak"Visual basic script"


elseif x="close" then

msgbox"Type stop to cancel the application! ",vbcritical,"Error interpreter.crash*$fg43@"
createobject("sapi.spvoice").speak"type stop to cancel the bot application"



elseif x="what is your name" then

createobject("sapi.spvoice").speak"hi, i'm microsoft anna   "




elseif x="py" then 

CreateObject("Wscript.shell").run"E:\py"




elseif x="hey anna fart for me" then

createobject("sapi.spvoice").speak"i have an audio clip of me farting, but unable to play due to technical issue with player"




elseif x="down" then 

CreateObject("Wscript.shell").run"C:\Users\Admin\Downloads"



elseif x="who is your gf" then

createobject("sapi.spvoice").speak"This is one of those things we'd both have to agree on. i'd prefer to keep our relationship friendly . you know, romance makes me incredibly awkward."



elseif x="who is your bf" then

createobject("sapi.spvoice").speak"that's a secret, i can't tell you now. Do not ask me why"



elseif x="sticky" then

CreateObject("Wscript.shell").run"%windir%\system32\StikyNot.exe"





elseif x="google drive" then

Createobject("wscript.shell").run"""https://drive.google.com/drive/my-drive"""




elseif x="political science" then

CreateObject("Wscript.shell").run"D:\Psci"




elseif x="ps notes" then

CreateObject("Wscript.shell").run"""D:\pol notes"""


elseif x="temp" then 

CreateObject("Wscript.shell").run"C:\Users\Admin\AppData\Local\Temp"




elseif x="recent" then 

CreateObject("Wscript.shell").run"C:\Users\Admin\Recent"









else


msgbox"sorry!! The interpreter did not undestand that.",vbcritical,"Error interpreter.crash*$fg43@"




createobject("sapi.spvoice").speak"sorry. i don't understand that. " 



end if





loop






