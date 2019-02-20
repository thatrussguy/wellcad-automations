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
' Calculate borehole deviation
obBorehole.CalculateBoreholeDeviation FALSE, deviationini
' Calculate borehole coordinates
obBorehole.CalculateBoreholeCoordinates FALSE, deviationini
' Calculate borehole closure
obBorehole.CalculateBoreholeClosure FALSE, deviationini
