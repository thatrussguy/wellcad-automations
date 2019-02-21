' Process deviation for Electromind OTV tool
' Tool requires magnetometer/inclinometer XYZ to be remapped to YZX
' ===========================================
' get active borehole to process
Set obWCAD=CreateObject("WellCAD.Application")
'obWCAD.ShowWindow()
Set obBorehole = obWCAD.GetActiveBorehole()
' ===========================================
' Set the folder details for the ini file
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
deviationini = strFolder &  "\Config files\Emind_OTV_deviation.ini"
' ===========================================
' Check if data was imported from a LAS file
' If so, depth shift all logs (except IMG AZ/IMG INCL) by 0.98m
Set obLog = obBorehole.Log("GX")
gxCreationInfo = obLog.HistoryItemDescription(1)
If InStr(1, gxCreationInfo, ".las", 1) > 0 Then
	depthshiftquery = msgbox("Looks like LAS file import. Shift depth?", 4, "Shift depth for LAS file?")
	If depthshiftquery = vbYes Then
		NbLogs = obBorehole.NbOfLogs
		For i = 1 to NbLogs
			If InStr(1, ObBorehole.Log(i - 1).Name, "IMG ", 1) < 1 Then
				obBorehole.DepthShiftLog i - 1, 0.98
			End If
		Next
	End If
End If
' ===========================================
' Calculate borehole deviation
obBorehole.CalculateBoreholeDeviation FALSE, deviationini
' Calculate borehole coordinates
obBorehole.CalculateBoreholeCoordinates FALSE, deviationini
' Calculate borehole closure
obBorehole.CalculateBoreholeClosure FALSE, deviationini
