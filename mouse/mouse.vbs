Set wshshell = CreateObject("wscript.shell")
wshshell.run "regsvr32  /n /i:user "".\wyhkm.dll"""

Set wyhkm=CreateObject("wyp.hkm")
ver = wyhkm.GetVersion()
MsgBox "Ä£¿é°æ±¾£º" & Hex(ver), 4096
