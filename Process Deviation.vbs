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
deviationini = strFolder &  "\Config files\deviation.ini"
' ===========================================
'Create functions
Function DeleteLog(curve)
	Set Borehole = obWCAD.GetActiveBorehole()
	If Not (Borehole.log(curve)) Is Nothing Then
		Borehole.RemoveLog curve , FALSE
	End If
End Function

Function ShiftLog(curve,shiftBy)
	Set Borehole = obWCAD.GetActiveBorehole()
	If Not (Borehole.log(curve)) Is Nothing Then
		Borehole.DepthShiftLog curve, shiftBy
	End If
End Function

Function RenameLog(oldname, newname)
	Set Borehole = obWCAD.GetActiveBorehole()
	Set obLog = Borehole.Log(oldname)
	obLog.Name = newname
End Function
' ===========================================
' Process deviation
' Electromind or Century?
' Abort if neither is chosen
toolTypeQuery = msgbox("Is this Electromind Deviation data?", 4, "Tool Type")
If toolTypeQuery = vbYes Then
' OTV or ATV data?
' Abort if neither is chosen
	imageTypeQuery = msgbox("Is this OTV Data?", 4, "Image Type")
	If imageTypeQuery = vbYes Then
' Delete logs
		DeleteLog "AZ"
		DeleteLog "IMG AZ"
		DeleteLog "IMG INCL"
		DeleteLog "INCL"
' LAS file import?
' Shift depth 0.98m if imported from OTV LAS file
	lasFileQuery = msgbox("Data imported from LAS file?", 4, "Las file?")
	If lasFileQuery = vbYes Then
		shiftDepth = 0.98
		ShiftLog "GX", shiftDepth
		ShiftLog "GY", shiftDepth
		ShiftLog "GZ", shiftDepth
		ShiftLog "HGDELTA", shiftDepth
		ShiftLog "HX", shiftDepth
		ShiftLog "HY", shiftDepth
		ShiftLog "HZ", shiftDepth
		ShiftLog "MROLL", shiftDepth
		ShiftLog "ROLL", shiftDepth
		ShiftLog "SPEED", shiftDepth
		ShiftLog "TGRAV", shiftDepth
		ShiftLog "TMAG", shiftDepth
	End If
' Calculate borehole deviation
		obBorehole.CalculateBoreholeDeviation FALSE, deviationini
' Calculate borehole coordinates
		obBorehole.CalculateBoreholeCoordinates FALSE, deviationini
' Calculate borehole closure
		obBorehole.CalculateBoreholeClosure FALSE, deviationini
' Delete raw deviation logs
		DeleteLog "GX"
		DeleteLog "GY"
		DeleteLog "GZ"
		DeleteLog "HX"
		DeleteLog "HY"
		DeleteLog "HZ"
	Else
		imageTypeQuery = msgbox("Is this ATV Data?", 4, "Image Type")
		If imageTypeQuery = vbYes Then
' Delete logs
			DeleteLog "AZ"
			DeleteLog "INCL"
' LAS file import?
' Shift depth 1.35m if imported from ATV LAS file
			lasFileQuery = msgbox("Data imported from LAS file?", 4, "Las file?")
			If lasFileQuery = vbYes Then
				shiftDepth = 1.35
				ShiftLog "GX", shiftDepth
				ShiftLog "GY", shiftDepth
				ShiftLog "GZ", shiftDepth
				ShiftLog "HGDELTA", shiftDepth
				ShiftLog "HX", shiftDepth
				ShiftLog "HY", shiftDepth
				ShiftLog "HZ", shiftDepth
				ShiftLog "MROLL", shiftDepth	
				ShiftLog "ROLL", shiftDepth
				ShiftLog "SPEED", shiftDepth
				ShiftLog "TGRAV", shiftDepth
				ShiftLog "TMAG", shiftDepth
			End If
' Calculate borehole deviation
			obBorehole.CalculateBoreholeDeviation FALSE, deviationini
' Calculate borehole coordinates
			obBorehole.CalculateBoreholeCoordinates FALSE, deviationini
' Calculate borehole closure
			obBorehole.CalculateBoreholeClosure FALSE, deviationini
' Delete raw deviation logs
			DeleteLog "GX"
			DeleteLog "GY"
			DeleteLog "GZ"
			DeleteLog "HX"
			DeleteLog "HY"
			DeleteLog "HZ"
			Else
				msgbox("Neither OTV or ATV image type selected. Deviation processing aborted.")
			End if
		End If
Else
	toolTypeQuery = msgbox("Is this Century Deviation data?", 4, "Tool Type")
	If toolTypeQuery = vbYes Then
' Delete logs
		DeleteLog "AZIMUTH"
		DeleteLog "D DIFF"
		DeleteLog "DISTANCE"
		DeleteLog "E DEV"
		DeleteLog "N DEV"
		DeleteLog "T DEPTH"
		DeleteLog "XFLUX"
		DeleteLog "XINCL"
		DeleteLog "YFLUX"
		DeleteLog "YINCL"
		DeleteLog "ZFLUX"
' Rename logs
		RenameLog "SANG", "Tilt"
		RenameLog "SANGB", "Azimuth"
' Calculate borehole coordinates
		obBorehole.CalculateBoreholeCoordinates FALSE, deviationini
' Calculate borehole closure
		obBorehole.CalculateBoreholeClosure FALSE, deviationini

	Else
		msgbox("Neither Electromind or Century tool type selected. Deviation processing aborted.")
	End If
End If
