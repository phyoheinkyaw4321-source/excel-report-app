import sys, time, subprocess, tempfile
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue

def prop(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

src = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2]).resolve()
profile = Path(tempfile.mkdtemp(prefix='lo_pivot_'))
proc = subprocess.Popen([
    'libreoffice','--headless',
    '--accept=socket,host=localhost,port=2011;urp;StarOffice.ComponentContext',
    f'-env:UserInstallation=file://{profile}'
], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    time.sleep(3)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = resolver.resolve('uno:socket,host=localhost,port=2011;urp;StarOffice.ComponentContext')
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(src)), '_blank', 0,
        (prop('Hidden', True), prop('UpdateDocMode', 3))
    )
    # Refresh every existing DataPilot/Pivot table. The template already contains
    # the required page filters, row field and Count - Task No value field.
    for sheet in doc.Sheets:
        dps = sheet.getDataPilotTables()
        for name in dps.getElementNames():
            dps.getByName(name).refresh()
    doc.calculateAll()
    if out.exists(): out.unlink()
    doc.storeAsURL(
        uno.systemPathToFileUrl(str(out)),
        (prop('FilterName', 'Calc MS Excel 2007 XML'), prop('Overwrite', True))
    )
    doc.close(True)
finally:
    proc.terminate()
    try: proc.wait(timeout=10)
    except Exception: proc.kill()
print(str(out))
