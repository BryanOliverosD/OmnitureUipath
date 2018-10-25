set objshell = createobject("wscript.shell")

Return = objshell.Run("taskkill /f /im iexplore.exe",0,False)

