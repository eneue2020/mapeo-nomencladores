Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = oWS.SpecialFolders("Desktop") & "\Mapeo Nomencladores.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "C:\Program Files\Python311\pythonw.exe"
oLink.Arguments = """C:\Users\Usuario\Desktop\Mapeo\app.py"""
oLink.WorkingDirectory = "C:\Users\Usuario\Desktop\Mapeo"
oLink.Description = "Mapeo de Nomencladores Medicos"
oLink.IconLocation = "C:\Program Files\Python311\pythonw.exe, 0"
oLink.Save
WScript.Echo "Acceso directo creado en el escritorio."
