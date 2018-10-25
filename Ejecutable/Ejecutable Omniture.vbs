set objshell = createobject("wscript.shell")

intReturn = objshell.Run("""C:\Program Files (x86)\UiPath\Studio\uirobot.exe"" /file""C:\Users\jcarrillo\Desktop\Omniture\Principal.xaml""" , 0, True)

