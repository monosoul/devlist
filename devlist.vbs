Dim strComputer : strComputer = "."
Dim objWMI : Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Dim colItems : Set colItems = objWMI.ExecQuery("Select * from Win32_PnPEntity WHERE Status = 'Error'", "WQL", WBEM_RETURN_IMMEDIATELY + WBEM_FORWARD_ONLY)
Dim objItem

For Each objItem in colItems
		Select Case objItem.ConfigManagerErrorCode
			Case 0
				errText = "Device is working properly."
			Case 1
				errText = "Device is not configured correctly."
			Case 2
				errText = "Windows cannot load the driver for this device."
			Case 3
				errText = "Driver for this device might be corrupted, or the system may be low on memory or other resources."
			Case 4
				errText = "Device is not working properly. One of its drivers or the registry might be corrupted."
			Case 5
				errText = "Driver for the device requires a resource that Windows cannot manage."
			Case 6
				errText = "Boot configuration for the device conflicts with other devices."
			Case 7
				errText = "Cannot filter."
			Case 8
				errText = "Driver loader for the device is missing."
			Case 9
				errText = "Device is not working properly. The controlling firmware is incorrectly reporting the resources for the device."
			Case 10
				errText = "Device cannot start."
			Case 11
				errText = "Device failed."
			Case 12
				errText = "Device cannot find enough free resources to use."
			Case 13
				errText = "Windows cannot verify the device's resources."
			Case 14
				errText = "Device cannot work properly until the computer is restarted."
			Case 15
				errText = "Device is not working properly due to a possible re-enumeration problem."
			Case 16
				errText = "Windows cannot identify all of the resources that the device uses."
			Case 17
				errText = "Device is requesting an unknown resource type."
			Case 18
				errText = "Device drivers must be reinstalled."
			Case 19
				errText = "Failure using the VxD loader."
			Case 20
				errText = "Registry might be corrupted."
			Case 21
				errText = "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device."
			Case 22
				errText = "Device is disabled."
			Case 23
				errText = "System failure. If changing the device driver is ineffective, see the hardware documentation."
			Case 24
				errText = "Device is not present, not working properly, or does not have all of its drivers installed."
			Case 25
				errText = "Windows is still setting up the device."
			Case 26
				errText = "Windows is still setting up the device."
			Case 27
				errText = "Device does not have valid log configuration."
			Case 28
				errText = "Device drivers are not installed."
			Case 29
				errText = "Device is disabled. The device firmware did not provide the required resources."
			Case 30
				errText = "Device is using an IRQ resource that another device is using."
			Case 31
				errText = "Device is not working properly. Windows cannot load the required device drivers."
		End Select
		MsgBox objItem.Name & vbCrLf & objItem.DeviceID & vbCrLf & errText
Next