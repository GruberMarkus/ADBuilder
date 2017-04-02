' ==============================================================================================
' AD-Builder.vbs
' ==============================================================================================
'
' Information about usage and input file format can be found in AD-Builder.txt.
'
' Changelog: AD-Builder_changelog.txt
'
' Author: markus.gruber@gruber.cc
'
' This script is open source and can be used for any project, private or commercial, without
' any fees.
'
' Please contact the author if you like this script, find a bug or have a general question.
'
' This script is provided "as is" and is used on your own risk.
'
' ==============================================================================================
' ==============================================================================================


' Display version information
VersionString = "v20090331.1300"
WScript.Echo "AD-Builder.vbs " & VersionString
WScript.Echo

' Script handles all errors on its own ...
On Error Resume Next

' Expand evironment variables in command strings: yes and when?
ExpandEnvironmentStringsBeforeRegExp = True
ExpandEnvironmentStringsAfterRegExp = True

' Allow cscript.exe only
If Not LCase(Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1)) = "cscript.exe" Then
	MsgBox "This script must be run with cscript.exe, not with wscript.exe." & VbCrLf & VbCrLf & "Exiting.", vbCritical + vbOKOnly, "Script started with wrong interpreter"
	WScript.Quit
End If

' Check number of passed arguments
Set objArgs = WScript.Arguments
If Not WScript.Arguments.count = 2 Then
	DisplayUsageInformationAndQuit
End If

' Check if correct arguments are passed, assign variables based on arguments
For Each strArg In objArgs
	If LCase(Left(strArg, Len("Inputfile="))) = LCase("Inputfile=") Then
		InputFile = Right(strArg, Len(strArg) - Len("Inputfile="))
	End If
	If LCase(Left(strArg, Len("SyntaxCheckOnly="))) = LCase("SyntaxCheckOnly=") Then
		ArgSyntaxCheckOnly = Right(strArg, Len(strArg) - Len("SyntaxCheckOnly="))
		If LCase(ArgSyntaxCheckOnly) = LCase("true") Then SyntaxCheckOnly = True
		If LCase(ArgSyntaxCheckOnly) = LCase("false") Then SyntaxCheckOnly = False
	End If
Next
If InputFile = "" Or SyntaxCheckOnly = "" Or Not ( SyntaxCheckOnly = True Or SyntaxCheckOnly = False ) Then
	DisplayUsageInformationAndQuit
End If

' Open input file and fill array line by line
Set objFSO = CreateObject("Scripting.FileSystemObject")
arrCommandLineArray = Split(objFSO.OpenTextFile(InputFile).ReadAll, vbNewLine)
If Err.Number <> 0 Then
	WScript.Echo "Error opening input file """ & InputFile & """. Exiting."
	WScript.Quit
End If

' Check for needed 3rd party files
Set WshShell = CreateObject("WScript.Shell")
Pathpath = WshShell.ExpandEnvironmentStrings("%path%") & ";" & CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
pathpatharray = Split(pathpath, ";")
FilesNeededlist = "dsadd.exe;dsacls.exe"
FilesNeeded = Split(FilesNeededList, ";")
AllNeededFilesFound = True
For Each y In FilesNeeded
	TempFileFound = False
	For Each x In pathpatharray
		If Right(x, 1) <> "\" Then x = x & "\"
		If TempFileFound = False Then
			If objFSO.FileExists(x & y) = True Then
				TempFileFound = True
			End If
		End If
	Next
	If TempFileFound = False Then
		If Missingfilelist = "" Then
			Missingfilelist = y
		Else
			Missingfilelist = Missingfilelist & ";" & y
		End If
		AllNeededFilesFound = False
	End If
Next
If AllNeededFilesFound = False Then
	WScript.echo "Not all needed files found."
	WScript.echo "Needed files: " & Replace(filesneededlist, ";", ", ") & "."
	WScript.echo "Not found: " & Replace(missingfilelist, ";", ", ") & "."
	WScript.echo "ERROR, exiting."
	WScript.quit
End If

' Check for installed GPMC
Set GPM = CreateObject("GPMgmt.GPM")
If Err.Number <> 0 Then
	WScript.echo "Line " & linenumberstring & ": GPO cannot be created. Is GPMC installed?"
	WScript.echo "Line " & linenumberstring & ": " & Err.number & ": " & Err.Description & "."
	WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
	WScript.quit
End If
Set GPM = Nothing

' Set some variables
AllowedCommandString = "ou;protectou;group;creategpo;linkgpo;perm"
CommandNameExecutionOrder = "ou;group;creategpo;linkgpo;perm"
Set WshShell = CreateObject("WScript.Shell")
Set VariableDictionary = CreateObject("Scripting.Dictionary")
ProtectOuFromAccidentalDeletionOnCreation = False
Err.Clear
Set objRootDSE = GetObject("LDAP://RootDSE")
strLDAPDomain = objRootDSE.Get("defaultNamingContext")
If Err.number <> 0 Then
	If SyntaxCheckOnly = False Then
		WScript.Echo "Error connecting to domain!"
		WScript.Echo "Error " & Err.number & ": " & Err.Description
		WScript.Echo "Exiting."
		WScript.Quit
	Else
		WScript.Echo "Error connecting to domain."
		WScript.Echo "Error " & Err.number & ": " & Err.Description
		WScript.Echo "Continuing with syntax check, but LDAP strings will show the wrong domain."
		WScript.Echo
	End If
End If
Err.Clear
If strLDAPDomain = "" Then strLDAPDomain = "dc=sample,dc=domain"
CurrentOU = "LDAP://" & strLDAPDomain
LineNumber = 0

' Inform user if SyntaxCheckOnly=true
If SyntaxCheckOnly = True Then
	WScript.Echo "SyntaxCheckOnly=True."
	WScript.Echo "The script will check the syntax of the input file, but not perform any changes."
	WScript.Echo
End If

' Take care of variables before doing anyting else
AllowedCommandArray = Split(AllowedCommandString, ";")
For i = LBound(arrCommandLineArray) To UBound(arrCommandLineArray)
	' Set LineNumber to correct value and format it to 4 digits with leading zeros
	LineNumber = i + 1
	Select Case Len(linenumber)
		Case 1
			linenumberstring = "000" & LineNumber
		Case 2
			linenumberstring = "00" & LineNumber
		Case 3
			linenumberstring = "0" & LineNumber
		Case 4
			linenumberstring = "" & LineNumber
	End Select
	strCommandLine = arrCommandLineArray(i)
	' remove whitespaces left and right of the string
	l = Left(strCommandLine, 1)
	r = Right(strCommandLine, 1)
	Do Until 0 = 7
		If l = vbTab Or l = " " Then
			strCommandLine = Mid(strCommandLine, 2)
			l = Left(strCommandLine, 1)
		ElseIf r = vbTab Or r = " " Then
			strCommandLine = Left(strCommandLine, Len(strCommandLine) - 1)
			r = Right(strCommandLine, 1)
		Else
			Exit Do
		End If
	Loop
	LineIsComment = 0
	LineIsCommand = 0
	If Left(strcommandline, 1) = "'" Or strCommandLine = "" Then
		'line is a comment or empty
		LineIsComment = 1
	Else
		'line should be an allowed command
		For Each AllowedCommand In AllowedCommandArray
			If InStr(LCase(strcommandline), LCase(AllowedCommand) & ": ") = 1 Then
				'line starts with an allowed command
				LineIsCommand = 1
			Else
				'line does not start with an allowed command, but there are several commands to check ...
				If LineIsCommand = 1 Then
					'there must be a prior match, do not change anything
				Else
					LineIsCommand = 0
				End If
			End If
		Next
	End If
	If LineIsCommand = 0 And LineIsComment = 0 Then
		'line is not a comment and not an allowed command
		If PossibleWrongSyntaxLineNumberList = "" Then
			PossibleWrongSyntaxLineNumberList = linenumberstring
		Else
			If Not (InStr(PossibleWrongSyntaxLineNumberList, linenumberstring) > 0) Then
				PossibleWrongSyntaxLineNumberList = PossibleWrongSyntaxLineNumberList & ";" & linenumberstring
			End If
		End If
	End If
	If LineIsCommand = 1 Then
		If ExpandEnvironmentStringsBeforeRegExp = True Then
			strCommandLine = WshShell.ExpandEnvironmentStrings(strCommandLine)
		End If
		' Use regular expressions
		' Find all variables, lower case them, write to array
		Set objRegEx = CreateObject("VBScript.RegExp")
		objRegEx.Global = True
		objRegEx.Pattern = "(\$Var[A-Za-z0-9]*\$)" 'A-Z, a-z and 0-9
		objRegEx.IgnoreCase = True
		Set TempString = objRegEx.Execute(strCommandLine)
		If TempString.count > 0 Then
			For Each expressionmatched In TempString
				' replace varibale names with lower case version
				strCommandLine = Replace(strCommandLine, expressionmatched, LCase(expressionmatched))
				arrCommandLineArray(i) = strCommandLine
				' Fill dictionary object with values
				If VariableDictionary.Exists(LCase(expressionmatched.value)) Then
					' variable is already there
				Else
					VariableDictionary.Add LCase(expressionmatched.value), LCase(expressionmatched.value)
				End If
			Next
		End If
		' If set, expand environment variables after working with regular expresseions
		If ExpandEnvironmentStringsAfterRegExp = True Then
			strCommandLine = WshShell.ExpandEnvironmentStrings(strCommandLine)
		End If
		' write back to array
		arrCommandLineArray(i) = strCommandLine
	Else
		arrCommandLineArray(i) = ""
	End If
Next
If PossibleWrongSyntaxLineNumberList <> "" Then
	WScript.Echo "Lines containing comment without leading ""'"" or wrong syntax, please check."
	WScript.Echo "================================================================================"
	WScript.Echo Replace(PossibleWrongSyntaxLineNumberList, ";", VbCrLf)
	WScript.Echo
	WScript.Echo
	WScript.Echo "ERROR, input file contains unclear statements. Exiting."
	WScript.Quit
End If

' Get values for the variables and replace the text in the input array
' Get Values
If VariableDictionary.Count > 0 Then
	KeysArray = VariableDictionary.Keys
	ItemsArray = VariableDictionary.Items
	Do Until UserAcceptsVariableValues = vbYes ' clicked on "yes"
		VariableSummary = ""
		For i = LBound(keysarray) To UBound(keysarray)
			VariableName = KeysArray(i)
			ItemsArray(i) = InputBox("Please enter the value for the variable:" & vbLf & KeysArray(i), "Enter value for variable", ItemsArray(i))
			If VariableSummary = "" Then
				VariableSummary = VariableName & ": " & ItemsArray(i)
			Else
				VariableSummary = VariableSummary & vbLf & VariableName & ": " & ItemsArray(i)
			End If
		Next
		UserAcceptsVariableValues = MsgBox(VariableSummary, vbYesNo + vbDefaultButton2 + vbQuestion, "Are the variable values correct?")
	Loop
	For i = LBound(arrCommandLineArray) To UBound(arrcommandlinearray)
		TempString = arrCommandLineArray(i)
		For j = LBound(keysarray) To UBound(keysarray)
			TempString = Replace(tempstring, LCase(KeysArray(j)), ItemsArray(j))
		Next
		arrCommandLineArray(i) = TempString
	Next
End If

'Let's go!
CommandNameExecutionOrderArray = Split(LCase(CommandNameExecutionOrder), ";")
For CommandNameExecutionOrderArrayID = LBound(CommandNameExecutionOrderArray) To UBound(CommandNameExecutionOrderArray)
	CommandNameExecutionOrderEntry = CommandNameExecutionOrderArray(CommandNameExecutionOrderArrayID)
	WScript.Echo "Taking care of command """ & CommandNameExecutionOrderEntry & """."
	WScript.Echo "================================================================================"
	LineNumber = 0
	' Parse file pasted into array line by line
	For Each strCommandLine In arrCommandLineArray
		' Set LineNumber to correct value and format it to 4 digits with leading zeros
		LineNumber = LineNumber + 1
		Select Case Len(linenumber)
			Case 1
				linenumberstring = "000" & LineNumber
			Case 2
				linenumberstring = "00" & LineNumber
			Case 3
				linenumberstring = "0" & LineNumber
			Case 4
				linenumberstring = "" & LineNumber
		End Select
		' Check if line is a comment or possibly has wrong syntax
		If strCommandLine = "" Then
			' do nothing, strCommandLine is empty
		Else
			' Line seems to be ok, go on by finding out which action to perform
			CommandTextRightPosition = InStr(1, strcommandline, ": ") - 1
			CommandTextLeftPosition = InStr(1, strcommandline, ": ") + 1
			CommandNameLCASE = LCase(Left(strcommandline, CommandTextRightPosition))
			' CommandName is the action to perform
			CommandName = Left(strcommandline, CommandTextRightPosition)
			'Now, the commands are sorted and forwarded to the correct sub functions
			Select Case CommandNameLCASE
				Case LCase("protectou")
					' The "protectou" switch must always be set when attempting to create OUs.
					If CommandNameExecutionOrderEntry = "ou" Then
						ProtectOUSwitch
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
				Case LCase("OU")
					' We always need "CurrentOU"
					strSearchString = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
					If InStr(strsearchstring, " -desc") > 0 Then
						temp2 = Split(strsearchstring, " -desc")
						temp = Split(temp2(0), "\")
						OUOnly = temp2(0)
					Else
						temp = Split(strsearchstring, "\")
						OUOnly = strSearchString
					End If
					stroualone = ""
					For i = UBound(temp) To 0 Step - 1
						If stroualone = "" Then
							stroualone = "ou=" & temp(i)
						Else
							stroualone = stroualone & ",ou=" & temp(i)
						End If
					Next
					CurrentOU = "LDAP://" & stroualone & "," & strLDAPDomain
					If CommandNameExecutionOrderEntry = "ou" Then
						CreateOU
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
				Case LCase("Group")
					If CommandNameExecutionOrderEntry = "group" Then
						CreateGroup
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
				Case LCase("perm")
					If CommandNameExecutionOrderEntry = "perm" Then
						SetPermission
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
				Case LCase("createGPO")
					If CommandNameExecutionOrderEntry = "creategpo" Then
						CreatePolicyInAD
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
				Case LCase("linkGPO")
					If CommandNameExecutionOrderEntry = "linkgpo" Then
						LinkPolicyToOu
						' Empty line for the log after each line from the array containing the file
						WScript.Echo
					End If
			End Select
		End If
	Next
	If CommandNameExecutionOrderArrayID < UBound(CommandNameExecutionOrderArray) Then
		WScript.Echo
		WScript.Echo
	End If
Next


Sub CreatePolicyInAD
	strGPO = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	strDomain = Replace(strLDAPdomain, "DC=", "")
	strDomain = Replace(strdomain, ",", ".")
	WScript.echo "Line " & linenumberstring & ": Creating GPO """ & strGPO & """."
	If SyntaxCheckOnly = False Then
		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand = CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Open "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection
		objCommand.Properties("Page Size") = 1000
		objCommand.CommandText = "<LDAP://cn=policies,cn=system," & strLDAPDomain & ">;(&(objectcategory=grouppolicycontainer)(objectclass=grouppolicycontainer)(displayname=" & strGPO & "));ADsPath;OneLevel"
		Set objRS = objCommand.Execute
		If objRS.EOF <> True Then
			objRS.MoveFirst
		End If
		If objRS.RecordCount > 0 Then
			' GPO with same name exists already
			WScript.echo "Line " & linenumberstring & ": PROBLEM, GPO """ & strGPO & """ already exists. Continuing with script."
		Else
			Set GPM = CreateObject("GPMgmt.GPM")
			Set Constants = GPM.GetConstants()
			Set GPMDomain = GPM.GetDomain(strDomain, "", Constants.UseAnyDC)
			On Error Resume Next
			Err.Clear
			Set GPMGPO = GPMDomain.CreateGPO()
			If Err.Number <> 0 Then
				WScript.echo "Line " & linenumberstring & ": Error creating GPO."
				WScript.echo "Line " & linenumberstring & ": " & Err.number & ": " & Err.Description & "."
				WScript.echo "Line " & linenumberstring & ": PROBLEM, continuing with script."
				Exit Sub
			End If
			' Now set the display name. If this fails, delete the GPO and report an error
			Err.Clear
			GPMGPO.DisplayName = strGPO
			If Err.Number <> 0 Then
				GPMGPO.Delete()
				WScript.echo "Line " & linenumberstring & ": Error creating GPO."
				WScript.echo "Line " & linenumberstring & ": " & Err.number & ": " & Err.Description & "."
				WScript.echo "Line " & linenumberstring & ": PROBLEM, continuing with script."
				Exit Sub
			End If
		End If
	End If
End Sub


Sub LinkPolicyToOu
	strGPO = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	strOuDN = Replace(currentou, "LDAP://", "")
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Provider = "ADsDSOObject"
	objConn.Open "Active Directory Provider"
	Set objRS = objConn.Execute("<LDAP://cn=policies,cn=system," & strLDAPDomain & ">;(&(objectcategory=grouppolicycontainer)(objectclass=grouppolicycontainer)(displayname=" & strGPO & "));adspath;OneLevel")
	If objRS.EOF <> True Then
		objRS.MoveFirst
	End If
	If objRS.RecordCount = 1 Then
		' GPO found
		strGPOADsPath = objRS.Fields(0).Value
	ElseIf objRS.RecordCount = 0 Then
		' GPO not found
		WScript.echo "Line " & linenumberstring & ": PROBLEM, GPO """ & strGPO & """ not found. Continuing with script."
		Exit Sub
	ElseIf objRS.RecordCount > 1 Then
		' More than 1 GPO found
		WScript.echo "Line " & linenumberstring & ": PROBLEM, found multiple GPOs named """ & strGPO & """. Continuing with script."
		Exit Sub
	End If
	WScript.Echo "Line " & linenumberstring & ": Linking GPO """ & strGPO & """ to """ & strOuDN & """."
	If SyntaxCheckOnly = False Then
		Set objOU = GetObject(CurrentOu)
		On Error Resume Next
		strGPLink = objOU.Get("gPLink")
		If InStr(LCase(strgplink), LCase(strgpoadspath)) > 0 Then
			' GPO is alread linked to that OU
			WScript.echo "Line " & linenumberstring & ": GPO already linked to OU."
		Else
			objOU.Put "gpLink", strGPLink & "[" & strGPOADsPath & ";0]"
			objOU.SetInfo
			If Not (Err.Number = 0 Or Err.number = "-2147463155") Then
				WScript.echo "Line " & linenumberstring & ": Error linking GPO."
				WScript.echo "Line " & linenumberstring & ": " & Err.number & ": " & Err.Description & "."
				WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
				WScript.Quit
			End If
		End If
	End If
End Sub


Sub ProtectOUSwitch
	strSearchString = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	Select Case strSearchString
		Case "on"
			WScript.Echo "Line " & linenumberstring & ": ""protectou"" set to ""on""."
			WScript.Echo "Line " & linenumberstring & ": Newly created OUs will be protected from accidental deletion."
			ProtectOuFromAccidentalDeletionOnCreation = True
		Case "off"
			WScript.Echo "Line " & linenumberstring & ": ""protectou"" set to ""off""."
			WScript.Echo "Line " & linenumberstring & ": Newly created OUs will not be protected from accidental deletion."
			ProtectOuFromAccidentalDeletionOnCreation = False
		Case Else
			WScript.Echo "Line " & linenumberstring & ": Wrong value for ""protectou"", only ""on"" or ""off"" is valid."
			WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
			WScript.Quit
	End Select
End Sub


Sub ProtectOuFromAccidentalDeletionOnCreationDo(ByVal CurrentOuDN)
	CurrentOuDN = Replace(CurrentOuDN, "LDAP://", "")
	ParentOuDN = Right(CurrentOuDN, ( Len(CurrentOuDN) - InStr(CurrentOuDN, ",") ))
	' Set permissions on parent OU
	Set oExec = WshShell.exec("%comspec% /c dsacls " & ParentOuDN & " /d everyone:DC")
	Do While Not oExec.StdOut.AtEndOfStream
		input = oExec.StdOut.Readall()
	Loop
	If oExec.ExitCode <> 0 Then
		WScript.Echo "Line " & linenumberstring & ": " & Replace(input, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
		WScript.echo "Line " & linenumberstring & ": Exit code """ & oExec.ExitCode & """."
		WScript.echo "Line " & linenumberstring & ": " & Replace(oExec.stderr.readall, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
		WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
		WScript.Quit
	Else
		WScript.Echo "Line " & linenumberstring & ": Success protecting parent OU """ & ParentOuDN & """."
	End If
	' Set permissions on current OU
	Set oExec = WshShell.exec("%comspec% /c dsacls " & CurrentOuDN & " /d everyone:SDDT")
	Do While Not oExec.StdOut.AtEndOfStream
		input = oExec.StdOut.Readall()
	Loop
	If oExec.ExitCode <> 0 Then
		WScript.Echo "Line " & linenumberstring & ": " & Replace(input, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
		WScript.echo "Line " & linenumberstring & ": Exit code """ & oExec.ExitCode & """."
		WScript.echo "Line " & linenumberstring & ": " & Replace(oExec.stderr.readall, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
		WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
		WScript.Quit
	Else
		WScript.Echo "Line " & linenumberstring & ": Success protecting newly created OU """ & CurrentOuDN & """."
	End If
End Sub


Sub CreateOU
	strSearchString = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	If InStr(strsearchstring, " -desc") > 0 Then
		temp2 = Split(strsearchstring, " -desc")
		temp = Split(temp2(0), "\")
		OUOnly = temp2(0)
	Else
		temp = Split(strsearchstring, "\")
		OUOnly = strSearchString
	End If
	stroualone = ""
	For i = UBound(temp) To 0 Step - 1
		If stroualone = "" Then
			stroualone = "ou=" & temp(i)
		Else
			stroualone = stroualone & ",ou=" & temp(i)
		End If
	Next
	OUDN = "LDAP://" & stroualone & "," & strLDAPDomain
	If InStr(strsearchstring, " -desc") > 0 Then
		ParameterString = Replace(temp2(0), OUONly, ( stroualone & "," & strLDAPDomain )) & " -desc" & temp2(1)
	Else
		ParameterString = Replace(strsearchstring, OUONly, ( stroualone & "," & strLDAPDomain ))
	End If
	CurrentOU = "LDAP://" & stroualone & "," & strLDAPDomain
	If ObjectExistsInAD(Replace(currentou, "LDAP://", ""), "organizationalUnit") = False Then
		If ProtectOuFromAccidentalDeletionOnCreation = True Then
			ProtectOUString = "protected"
		Else
			ProtectOUString = "not protected"
		End If
		WScript.echo "Line " & linenumberstring & ": Creating OU """ & stroualone & "," & strLDAPDomain & """, " & ProtectOUString & "."
		WScript.Echo "Line " & linenumberstring & ": " & "dsadd ou " & ParameterString
		If SyntaxCheckOnly = False Then
			Set oExec = WshShell.exec("%comspec% /c dsadd ou " & ParameterString)
			Do While Not oExec.StdOut.AtEndOfStream
				input = oExec.StdOut.Readall()
			Loop
			WScript.Echo "Line " & linenumberstring & ": " & Replace(input, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
			If oExec.ExitCode <> 0 Then
				WScript.echo "Line " & linenumberstring & ": Exit code """ & oExec.ExitCode & """."
				WScript.echo "Line " & linenumberstring & ": " & Replace(oExec.stderr.readall, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
				WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
				WScript.Quit
			End If
			If ProtectOuFromAccidentalDeletionOnCreation = True Then ProtectOuFromAccidentalDeletionOnCreationDo(CurrentOU)
		End If
	Else
		' CurrentOU is already set to an existing value
		WScript.echo "Line " & linenumberstring & ": OU """ & stroualone & "," & strLDAPDomain & """ already exists."
	End If
End Sub


Sub CreateGroup
	strSearchString = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	If InStr(strsearchstring, " ") > 0 Then
		temp = Split(strsearchstring, " ")
		GroupName = temp(0)
	Else
		GroupName = strSearchString
	End If
	fullgrouppath = "cn=" & GroupName & "," & Replace(currentou, "LDAP://", "")
	ParameterString = Replace(strsearchstring, groupname, fullgrouppath)
	If ObjectExistsInAD(fullgrouppath, "group") = False Then
		WScript.echo "Line " & linenumberstring & ": Creating group """ & GroupName & """ in """ & Replace(currentou, "LDAP://", "") & """."
		WScript.Echo "Line " & linenumberstring & ": " & "dsadd group " & ParameterString
		If SyntaxCheckOnly = False Then
			Set oExec = WshShell.exec("%comspec% /c dsadd group " & ParameterString)
			Do While Not oExec.StdOut.AtEndOfStream
				input = oExec.StdOut.Readall()
			Loop
			WScript.Echo "Line " & linenumberstring & ": " & Replace(input, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
			If oExec.ExitCode <> 0 Then
				WScript.echo "Line " & linenumberstring & ": Exit code """ & oExec.ExitCode & """."
				WScript.echo "Line " & linenumberstring & ": " & Replace(oExec.stderr.readall, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
				WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
				WScript.Quit
			End If
		End If
	Else
		WScript.echo "Line " & linenumberstring & ": group """ & fullgrouppath & """ already exists."
	End If
End Sub


Sub SetPermission
	strSearchString = Right(strcommandline, Len(strcommandline) - CommandTextLeftPosition)
	ParameterString = strSearchString
	WScript.echo "Line " & linenumberstring & ": Setting permission """ & ParameterString & """ on """ & Replace(currentou, "LDAP://", "") & """."
	WScript.Echo "Line " & linenumberstring & ": " & "dsacls " & Replace(currentou, "LDAP://", "") & " " & ParameterString
	If SyntaxCheckOnly = False Then
		SetPermissionSuccess = 0
		SetPermissionRetryCount = 1
		SetPermissionRetryStop = 0
		Do Until SetPermissionRetryStop = 1
			WScript.echo "Line " & linenumberstring & ": Try " & SetPermissionRetryCount & " of 3."
			Set oExec = WshShell.exec("%comspec% /c dsacls " & Replace(currentou, "LDAP://", "") & " " & ParameterString)
			Do While Not oExec.StdOut.AtEndOfStream
				output = oExec.StdOut.Readall()
			Loop
			If oExec.ExitCode <> 0 Then
				WScript.echo "Line " & linenumberstring & ": Try " & SetPermissionRetryCount & " of 3 failed."
				WScript.Echo "Line " & linenumberstring & ": " & Replace(output, VbCrLf, VbCrLf & "Line " & linenumberstring & ": ")
				WScript.echo "Line " & linenumberstring & ": Exit code """ & oExec.ExitCode & """."
				SetPermissionRetryCount = SetPermissionRetryCount + 1
				SetPermissionSleepSeconds = SetPermissionRetryCount * 5
				If SetPermissionRetryCount < 4 Then
					WScript.echo "Line " & linenumberstring & ": Sleeping for " & SetPermissionSleepSeconds & " seconds."
					WScript.Sleep(SetPermissionSleepSeconds * 1000)
				End If
			Else
				SetPermissionSuccess = 1
				SetPermissionRetryCount = SetPermissionRetryCount + 1
				SetPermissionRetryStop = 1
				WScript.echo "Line " & linenumberstring & ": The command completed successfully."
			End If
			If SetPermissionRetryCount = 4 Then SetPermissionRetryStop = 1
		Loop
		If SetPermissionRetryCount = 4 And SetPermissionSuccess = 0 Then
			WScript.echo "Line " & linenumberstring & ": ERROR, exiting."
			WScript.Quit
		End If
	End If
End Sub


Function ObjectExistsInAD(ByVal ObjectExistsInADDN, ByVal ObjectExistsInADCategory)
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 1000
	objCommand.CommandText = "<LDAP://" & strLDAPDomain & ">;(objectCategory=" & ObjectExistsInADCategory & ");distinguishedName;Subtree"
	Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst
	ObjectExistsInAD = False
	If LCase(ObjectExistsInADDN) = LCase(Replace(strLDAPdomain, "LDAP://", "")) Then
		ObjectExistsInAD = True
	Else
		Do Until objRecordSet.EOF
			If ObjectExistsinAd = True Then Exit Do
			If LCase(objRecordSet.Fields("distinguishedName").Value) = LCase(ObjectExistsinADDN) Then
				ObjectExistsInAD = True
			End If
			objRecordSet.MoveNext
		Loop
	End If
	ObjectExistsInADDN = ""
	ObjectExistsInADCategory = ""
End Function


Sub DisplayUsageInformationAndQuit
	WScript.Echo "Error!"
	WScript.Echo "Wrong number of arguments passed or arguments have invalid values."
	WScript.Echo
	WScript.Echo "Correct syntax:"
	WScript.Echo "cscript.exe AD-Builder.vbs Inputfile=<path to file> SyntaxCheckOnly=[True|False] //nologo"
	WScript.Echo
	WScript.Echo "For best results and documention, pipe the output to a text file:"
	WScript.Echo "cscript.exe AD-Builder.vbs Inputfile=<path to file> SyntaxCheckOnly=[True|False] //nologo > <path to output file>"
	WScript.Quit
End Sub