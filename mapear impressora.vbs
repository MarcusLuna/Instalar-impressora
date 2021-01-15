'Instalando Porta TCP/IP da Impressora

Set objWMIService = GetObject("winmgmts:")
Set objNewPort = objWMIService.Get _
("Win32_TCPIPPrinterPort").SpawnInstance_
objNewPort.Name = "IP_192.168.68.2"
objNewPort.Protocol = 1
objNewPort.HostAddress = "192.168.68.2"
objNewPort.PortNumber = "9100"
objNewPort.SNMPEnabled = False
objNewPort.Put_


'Mapeando o Dispositivo de Impressão HP CP 1515n

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_

objPrinter.DriverName = "Kyocera ECOSYS M2040dn KX"
objPrinter.PortName = "IP_192.168.68.2"
objPrinter.DeviceID = "Diretoria"
objPrinter.Location = "Sala das Secretárias Executivas / COlor Laser"
objPrinter.Network = True
objPrinter.Put_

'Definindo o Dispositivo de Impressão Padrão
Set colInstalledPrinters = objWMIService.ExecQuery _
("Select * from Win32_Printer Where Name = 'Diretoria'")

For Each objPrinter in colInstalledPrinters
objPrinter.SetDefaultPrinter()
objPrinter.Put_

Next