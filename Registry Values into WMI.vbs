
'================================================================
'
' Version:	2.0
'
' Revision: Original Release - Victor Meyer (victor.meyer@salt.ky) (05/01/2015)
'			Updated Release - (victor.meyer@salt.ky) (26/01/2015)
'
' Versions: See https://github.com/VJayMeyer/Collect-Registry-Keys-with-Configuration-Manager
'
' Purpose: 	This script will loop through the comma seperated list of registry
'			paths and store the values within the CM_RegistryValues WMI Class
'			
' Inputs:	The comma seperated array labeled KEY_PATHS bellow in the following
'			format. "HKEY_LOCAL_MACHINE\Path1","HKEY_LOCAL_MACHINE\Path2"
'			
' Notes:	All registry key types will be converted to strings for WMI Storage
'
' Warning:	This script will collect all sub keys under the KEY_PATH and as such
'			Should the path be too shallow this will result in HINV (SCCM Inventory)
'			failures because of inventory file sizes. Increases to the "Max MIF Size"
'			Should be carefully considered and follow standard testing and change 
'			management processes.
'
' Outputs:	An up to date CM_RegistryValues WMI class which will be collected via 
'				SCCM Hardware Inventory
'			An event in the event log stating the start and end of the collection
'				as well as any error codes and messages. Look for Source = WSH
'
'================================================================
' ENTER KEY PATHS BELOW AS PER THE INSTRUCTIONS ABOVE
'================================================================
KEY_PATHS = Array("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations", _
 					"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\LAB")
'================================================================
' 				EDIT BELOW THIS LINE WITH CAUTION!
'================================================================
' True = Debug Printing On | False = Debug Printing Off
DEBUG_PRINTING = False 
' Call Master Execution Sub
MASTER_EXECUTION 
'================================================================
' 				SUBS AND FUNCTIONS
'================================================================
Sub MASTER_EXECUTION()

	' WMI Class Management
	MAINTAIN_WMI_CLASS()
	
	' Registry Key Storage
	For Each KEY_PATH In KEY_PATHS
		STORE_KEYS(KEY_PATH)
	Next
	
	' Log Completion Status
	If Err.Number <> 0 Then
		EVENT_WRITER "ERROR","Storing Registry Keys Failed " & _
			Err.Number & " | " & Err.Description
	Else
		EVENT_WRITER "INFO","Storing Registry Keys Completed Successfully"					
	End If

End Sub

Function CONVERT_HIVE(HIVE)

	' Purpose: 	This Function will covert a hive from its friendly name
	'			to systems name.
	'			
	' Inputs:	A Hive in the friendly name format
	'				Example: HKEY_LOCAL_MACHINE
	'
	' Outputs:	This will return a hives system name
	'				Example: HKEY_LOCAL_MACHINE = &H80000002
	'
	' Limits:	This will only currently work with the following Hives:-
	'				HKEY_LOCAL_MACHINE 	= &H80000002
	'				HKEY_USERS 			= &H80000003
	'				HKEY_CURRENT_CONFIG	= &H80000005

	' Check and return a system name based on a friendly name
	If UCase(HIVE) = "HKEY_LOCAL_MACHINE" Then
		' Return the System Name
		CONVERT_HIVE = &H80000002
	
	ElseIf UCase(HIVE) = "HKEY_USERS" Then
		' Return the System Name
		CONVERT_HIVE = &H80000002
	
	ElseIf UCase(HIVE) = "HKEY_CURRENT_CONFIG" Then
		' Return the System Name
		CONVERT_HIVE = &H80000005
	
	Else
		
		' Write to event log
		EVENT_WRITER "ERROR","Converting Hive " & _
			HIVE & " failed - " & Err.Number & " | " & Err.Description
		' Quit
		WScript.Quit
	
	End If

End Function

Sub STORE_KEYS(KEY_PATH)
	
	' Set the Process Key Flag to False (Start Position)
	blnProcess = False
	
	' Parse the Key Path
	If InStr(KEY_PATH,"HKEY_LOCAL_MACHINE") Then
		' Break the Key Path Down
		arrPieces = Split(KEY_PATH,"HKEY_LOCAL_MACHINE")
		If UBound(arrPieces) = 1 Then
			HIVE = "HKEY_LOCAL_MACHINE"
			KEY = Right(arrPieces(1),Len(arrPieces(1))-1)	
			blnProcess = True	
		Else
			blnProcess = False
		End If		
	ElseIf InStr(KEY_PATH,"HKEY_USERS") Then
		' Break the Key Path Down
		arrPieces = Split(KEY_PATH,"HKEY_USERS")
		If UBound(arrPieces) = 1 Then
			HIVE = "HKEY_USERS"
			KEY = Right(arrPieces(1),Len(arrPieces(1))-1)			
			blnProcess = True
		Else
			blnProcess = False
		End If	
	ElseIf InStr(KEY_PATH,"HKEY_CURRENT_CONFIG") Then
		' Break the Key Path Down
		arrPieces = Split(KEY_PATH,"HKEY_CURRENT_CONFIG")
		If UBound(arrPieces) = 1 Then
			HIVE = "HKEY_CURRENT_CONFIG"
			KEY = Right(arrPieces(1),Len(arrPieces(1))-1)
			blnProcess = True
		Else
			blnProcess = False
		End If
	
	Else
		blnProcess = False		
	End If 
	
	' Determine whether key processing should start
	If blnProcess = False Then ' Skip the registry key		
		' Write a warning and abort the value load
		EVENT_WRITER "WARNING","The registry path: " & KEY_PATH & _
				" is malformed or contains" & _
				" a hive that is not accessible" & vbCrLf & vbCrLf & _
				" This key will not be processed further"	
	Else ' Perform the Registry load. 			
		' Check that the Key_Path is not pointing directly to a value
		If IsRegistryValue(KEY_PATH) = "True" Then	
		
			' Enumerate the key as a path directly to a value
			If DEBUG_PRINTING = True Then
				WScript.Echo "****************************************************************************************"		
				WScript.Echo "Reading Single Value For: " & HIVE & "\" & KEY
				WScript.Echo "****************************************************************************************"
			End If
			' Pass the Hive and Key to the value reader.
			VALUE_READER HIVE, Replace(KEY,ValueFinder(KEY),"") ,ValueFinder(KEY)	
			
			If Err.Number <> 0 Then
				EVENT_WRITER "ERROR","Enumerating registry value " & _
					KEY_PATH & " failed " & Err.Number & " | " & Err.Description
			Else
				EVENT_WRITER "INFO","Enumerating registry value " & _
					KEY_PATH & " completed successfully"	
			End If
			On Error Goto 0			

		Else ' Its a Key

			' Enumerate all values in the root key
			If DEBUG_PRINTING = True Then
				WScript.Echo "****************************************************************************************"		
				WScript.Echo "Reading Root Key: " & HIVE & "\" & KEY
				WScript.Echo "****************************************************************************************"
			End If
			VALUE_READER HIVE, KEY, ""
			
			' Enumerate all keys
			SUBKEY_READER HIVE, KEY
			
			If Err.Number <> 0 Then
				EVENT_WRITER "ERROR","Enumerating registry key " & _
					KEY_PATH & " failed " & Err.Number & " | " & Err.Description
			Else
				EVENT_WRITER "INFO","Enumerating registry key " & _
					KEY_PATH & " completed successfully"	
			End If
			On Error Goto 0

		End If

 	End If
  
End Sub 

Function ValueFinder(PATH)

	ValueFinder = Right(PATH,Len(PATH) - InStrRev(PATH,"\"))

End Function

Function VALUE_READER(HIVE, KEY, SEARCHFILTER)

	' Bind to the WMI Registry Object	
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		objRegistry.EnumValues CONVERT_HIVE(HIVE), KEY,arrValueNames, arrValueTypes
	Set objRegistry = Nothing
	If IsNull(arrValueNames) = False Then
		' Loop through values
		For I=0 To UBound(arrValueNames)
			' Check there is a value name
			If LEN(arrValueNames(I)) >= 1 Then
				' Check if there is a SearchFilter and use it
				If Len(SEARCHFILTER) >= 1 Then
					If arrValueNames(I) = SEARCHFILTER Then
						VALUE_DT_READER HIVE,KEY,arrValueTypes(I),arrValueNames(I)
				    End If
				Else
					VALUE_DT_READER HIVE,KEY,arrValueTypes(I),arrValueNames(I)
				End If
			End If
		Next
	End If

	' Close down any unused objects
	Set objRegistry = Nothing

End Function

SUB VALUE_DT_READER(HIVE,KEY,VALUE_TYPE,VALUE_NAME)

	' Registry Data Dype Constants
	const REG_SZ = 1
	const REG_EXPAND_SZ = 2
	const REG_BINARY = 3	
	const REG_DWORD = 4
	const REG_MULTI_SZ = 7

	' Connect to the Registry
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

	' Run through possible data types and store values
	Select Case VALUE_TYPE
	Case REG_SZ
	    ' Get String Value"
	    objRegistry.GetStringValue CONVERT_HIVE(HIVE),KEY,VALUE_NAME,strValue
	    If DEBUG_PRINTING = True Then
	    	WScript.Echo "Value for key: " & KEY & "\" & VALUE_NAME & " of type: REG_SZ Is: " & CStr(strValue)
	    End If
	    ' Write String Value and Associated Data to WMI
	    WMI_WRITER VALUE_NAME,strValue,"REG_SZ",HIVE & "\" & KEY & "\" & VALUE_NAME
	Case REG_EXPAND_SZ
	    ' Get Expanded String value
	    objRegistry.GetExpandedStringValue CONVERT_HIVE(HIVE),KEY,VALUE_NAME,strValue
	    If DEBUG_PRINTING = True Then
	    	WScript.Echo "Value for key: " & KEY & "\" & VALUE_NAME & " of type: REG_EXPAND_SZ Is: " & strValue
	    End If
	    ' Write Expanded String Value and Associated Data to WMI
	    WMI_WRITER VALUE_NAME,strValue,"REG_EXPAND_SZ",HIVE & "\" & KEY & "\" & VALUE_NAME
	    
	Case REG_BINARY
		' Get Binary Value
	    objRegistry.GetBinaryValue CONVERT_HIVE(HIVE),KEY,VALUE_NAME,arrValues
	    If DEBUG_PRINTING = True Then
	    	WScript.Echo "Value for key: " & KEY & "\" & VALUE_NAME & " of type: REG_BINARY Is: " & BINARY_TO_STRING(arrValues)
	    End If
	    ' Write Binary Value and Associated Data to WMI
	    WMI_WRITER VALUE_NAME,BINARY_TO_STRING(arrValues),"REG_BINARY",HIVE & "\" & KEY & "\" & VALUE_NAME
	    
	Case REG_DWORD
	    ' Get DWORD Value
	    objRegistry.GetDWORDValue CONVERT_HIVE(HIVE),KEY,VALUE_NAME,strValue
	    If DEBUG_PRINTING = True Then
	    	WScript.Echo "Value for key: " & KEY & "\" & VALUE_NAME & " of type: REG_DWORD Is: " & strValue		            	
	    End If
	    ' Write DWORD Value and Associated Data to WMI
		WMI_WRITER VALUE_NAME,strValue,"REG_DWORD",HIVE & "\" & KEY & "\" & VALUE_NAME		

    Case REG_MULTI_SZ
        ' Get Multi String Value"         
        Return = objRegistry.GetMultiStringValue(CONVERT_HIVE(HIVE),KEY,VALUE_NAME,arrValues)
		If (Return = 0) Then
			If DEBUG_PRINTING = True Then		
				WScript.Echo "Value for key: " & KEY & "\" & VALUE_NAME &  " of type: REG_MULTI_SZ Is: " & ARRAY_TO_STRING(arrValues)
			End If
			' Write Multi String" Value and Associated Data to WMI
			WMI_WRITER VALUE_NAME,ARRAY_TO_STRING(arrValues),"REG_MULTI_SZ",HIVE & "\" & KEY & "\" & VALUE_NAME
		Else
			If DEBUG_PRINTING = True Then
		    	Wscript.Echo "GetMultiStringValue failed for key" & KEY & "\" & VALUE_NAME & ". Error = " & Err.Number & " | " & Err.Description
		    End If
		    EVENT_WRITER "ERROR","GetMultiStringValue failed for " & VALUE_NAME & _
					HIVE & "\" & KEY & " failed " & Err.Number & " | " & Err.Description
		End If
	End Select

	Set objRegistry = Nothing

End SUB

Function SUBKEY_READER(HIVE, KEY)

	If DEBUG_PRINTING = True Then
		WScript.Echo "****************************************************************************************"
		WScript.Echo "Reading Subkey: " & HIVE & "\" & KEY
		WScript.Echo "****************************************************************************************"
	End If

	' Bind to the WMI Registry Object	
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	' Enumerate Sub Keys
	objRegistry.EnumKey CONVERT_HIVE(HIVE), KEY, SUBKEYS
	' Ensure there are in fact subkeys
	If Not IsNull(SUBKEYS) Then
		For Each SUBKEY In SUBKEYS
    		' Load values for the subkey
    		VALUE_READER HIVE, KEY & "\" & SUBKEY, ""
    		' Check for further sub keys
      		SUBKEY_READER HIVE, KEY & "\" & SUBKEY    		
		Next
	Else
		' Just read the values at the key level
		VALUE_READER HIVE, KEY, ""
	End If
	
	' Close down any unused objects
	Set objRegistry = Nothing

End Function

Sub MAINTAIN_WMI_CLASS()

	'
	' Declare WMI Column Data Types
	wbemCimtypeSint16 = 	2 
	wbemCimtypeSint32 = 	3 
	wbemCimtypeReal32 = 	4 
	wbemCimtypeReal64 = 	5 
	wbemCimtypeString = 	8 
	wbemCimtypeBoolean = 	11 
	wbemCimtypeObject = 	13 
	wbemCimtypeSint8 = 	16 
	wbemCimtypeUint8 = 	17 
	wbemCimtypeUint16 = 	18 
	wbemCimtypeUint32 = 	19 
	wbemCimtypeSint64 = 	20 
	wbemCimtypeUint64 = 	21 
	wbemCimtypeDateTime = 	101 
	wbemCimtypeReference = 	102 
	wbemCimtypeChar16 = 	103 

	' Setup WMI Connection Objects
	Set objWMILocation = CreateObject("WbemScripting.SWbemLocator") 
	
	' Connect to the root\cimv2 Class
	Set objWMIServices = objWMILocation.ConnectServer(,"root\cimv2")
	
	' CM_RegistryValues existance trap
	On Error Resume Next
	
	' Connect to the CM_RegistryValues Class
	Set CM_REGISTRYVALUES = objWMIServices.Get("CM_RegistryValues")
	
	' Trap the error number for class missing
	If Err.Number = -2147217406 Then
		' CM_REGISTRYVALUES Class does not exist
		EVENT_WRITER "INFO","Initial CM_RegistryValues Class Creation detected"
	ElseIf Err.Number = 0 Then  
		' Delete the CM_REGISTRYVALUES Class
		CM_REGISTRYVALUES.Delete_ 
		EVENT_WRITER "INFO","CM_RegistryValues Class Deleted Successfully"
	Else
		' Fatal error
		EVENT_WRITER "ERROR","Fatal MAINTAIN_WMI_CLASS Function failure: " & Err.Number & " | " & Err.Number
		WScript.Quit
	End If 
	On Error Goto 0	

	On Error Resume Next
	' Create the CM_RegistryValues Class
	Set objWMI = objWMIServices.Get 
		objWMI.Path_.Class = "CM_RegistryValues" 
		objWMI.Properties_.add "KeyPath" , wbemCimtypeString
		objWMI.Properties_("KeyPath").Qualifiers_.add "key" , True
		objWMI.Properties_.add "Name" , wbemCimtypeString 
		objWMI.Properties_.add "Value" , wbemCimtypeString
		objWMI.Properties_.add "DataType" , wbemCimtypeString		

		objWMI.Put_ 
				
	If Err.Number = 0 Then   
		EVENT_WRITER "INFO","CM_RegistryValues Class Created Successfully"
	Else
		' Fatal error
		EVENT_WRITER "ERROR","Fatal MAINTAIN_WMI_CLASS Function failure: " & Err.Number & " | " & Err.Description
		WScript.Quit
	End If
	On Error Goto 0
		
	' Close down any unused objects
	Set objWMILocation = Nothing
	Set objWMIServices = Nothing
	Set objWMI = Nothing

End Sub

Sub EVENT_WRITER(EVENT_TYPE,EVENT_MSG)

	' Create the Wscript.Shell object
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	' Parse the Event Type and Write to the Event Log
	If EVENT_TYPE = "WARNING" Then
		objShell.LogEvent 2, EVENT_MSG		
	ElseIf EVENT_TYPE = "ERROR" Then		
		objShell.LogEvent 1, EVENT_MSG
	Else
		objShell.LogEvent 4, EVENT_MSG
	End If
	
	' Close down any unused objects
	Set objShell = Nothing

End Sub

Function WMI_WRITER(NAME,VALUE,DATATYPE,KEYPATH)
	
		' Check that all variables have a name value (key)
		' Setup WMI Connection Object
		Set objWMILocation = CreateObject("WbemScripting.SWbemLocator") 
		
		' Connect to the root\cimv2 Class
		Set objWMIServices = objWMILocation.ConnectServer(,"root\cimv2")
	
		' Instantiate a new 
		Set objWMI = objWMIServices.Get("CM_RegistryValues" ).SpawnInstance_ 

		If DEBUG_PRINTING = True Then
			WScript.Echo "==========================================================="
			WScript.Echo "WMI WRITE - NAME: " & NAME
			WScript.Echo "WMI WRITE - VALUE: " & VALUE
			WScript.Echo "WMI WRITE - DATATYPE: " & DATATYPE
			WScript.Echo "WMI WRITE - KEYPATH: " & KEYPATH
		End If
	
		' Populate Values
		objWMI.Name = NAME
		objWMI.Value = VALUE
		objWMI.DataType = DATATYPE
		objWMI.KeyPath = KEYPATH
		
		' Commit Values
		On Error Goto 0 
		objWMI.Put_
		
		If DEBUG_PRINTING = True Then
			WScript.Echo "Error Details After Write: " & Err.Number & " | " & Err.Description
			WScript.Echo "==========================================================="
		End If
	
		' Close down any unused objects
		Set objWMILocation = Nothing
		Set objWMIServices = Nothing
		Set objWMI = Nothing
		
End Function

Function ARRAY_TO_STRING(arrayname)
	' Strip and comma seperate array
	If Ubound(arrayname) >= 0 Then
		strText = CStr(arrayname(LBound(arrayname)))
		for i=LBound(arrayname)+1 to UBound(arrayname)
			strText = strText & "," & CStr(arrayname(i))
		next
		ARRAY_TO_STRING = strText
	End If
End Function

Function BINARY_TO_STRING(arrValues)
	' Strip and convert binary data to string
	strText = strText & ": "
	For Each strValue in arrValues
		strText = strText & " " & strValue 
	Next 
	BINARY_TO_STRING = strText
End function 

Function IsRegistryValue (strRegistryKey)
	On Error Resume Next
   
    Set objShell = CreateObject("WScript.Shell")
    value = objShell.RegRead( strRegistryKey )

    If err.number <> 0 Then
		IsRegistryValue = "False"
    Else
		IsRegistryValue = "True"
    End If
    
	On Error Goto 0
	
    Set WSHShell = nothing
End Function
