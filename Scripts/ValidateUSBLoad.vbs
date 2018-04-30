Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

bAllClear = True
'For unit testing this should be false, for first unit true then fill in the new values below. 
bFirstUnitTest = False

'Change these values reflect the files you want to test, update code on lines 38-40
sVolName = ""
sExeMD5 = ""
sPdfMd5 = ""

Count = 0

'Get a list of drives to check
Set colDrives = objFSO.Drives
For Each objDrive in colDrives
	'Change/add the drive letter exclude for the computer that you are running on.
	if objDrive.DriveLetter <> "C" THEN
		For Each objItem in objWMIService.ExecQuery ("Select * From Win32_LogicalDisk Where DeviceID = '" & objDrive.DriveLetter & ":'")	
			'Files to check 
			Count = Count + 1
			'update to match lines 9-11
			sFile_1 = objDrive.DriveLetter & ":\[TestFile].exe"
			sFile_2 = objDrive.DriveLetter & ":\[TestFile].pdf"
			'Set the default to blank in case the file is not found 
			sHash_1 = ""
			sHash_2 = ""
			'Check if the file exists before testing it 
			IF objFSO.FileExists(sFile_1) Then
				sHash_1 = BytesToBase64(md5hashBytes(GetBytes(sFile_1)))
			END IF 
			IF objFSO.FileExists(sFile_1) Then
				sHash_2 = BytesToBase64(md5hashBytes(GetBytes(sFile_2)))
			END IF
			'For the first run unit to get the values for lines 9-11. 
			IF bFirstUnitTest THEN 
				Wscript.Echo "sVolName = " & CHR(34) & objItem.VolumeName  & CHR(34)
				Wscript.Echo "sExeMD5 = "  & CHR(34) & sHash_1  & CHR(34)
				Wscript.Echo "sPdfMd5 = "  & CHR(34) & sHash_2  & CHR(34)
			ELSE 
				'If the values do not match then output the error on screen, update to match line 9-11
				IF  objItem.VolumeName <> sVolName OR sExeMD5 <> sHash_1 OR sPdfMd5 <> sHash_2 THEN 
					Wscript.Echo "Drive letter: " & objDrive.DriveLetter
					Wscript.Echo "Volume Name: " & objItem.VolumeName
					Wscript.Echo "EXE: " & sHash_1
					Wscript.Echo "PDF: " & sHash_2
					Wscript.Echo " "
					bAllClear = false 
				END IF 
			END IF 
		Next
	END IF
Next

'Output if their was an issue or not. 
IF NOT bFirstUnitTest THEN 
	IF NOT bAllClear THEN 
		Wscript.Echo "Issues were found."
	ELSE 
		Wscript.Echo "All units passed: " & Count
	END IF 
END IF 

Wscript.Echo " "
	
'MD5 code from http://stackoverflow.com/questions/10198690/how-to-generate-md5-using-vbscript-in-classic-asp
function md5hashBytes(aBytes)
    Dim MD5
    set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    MD5.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    md5hashBytes = MD5.ComputeHash_2( (aBytes) )
end function

function sha1hashBytes(aBytes)
    Dim sha1
    set sha1 = CreateObject("System.Security.Cryptography.SHA1Managed")

    sha1.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha1hashBytes = sha1.ComputeHash_2( (aBytes) )
end function

function sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha256hashBytes = sha256.ComputeHash_2( (aBytes) )
end function

function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function

function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function

Function BytesToBase64(varBytes)
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = .Text
    End With
End Function

Function GetBytes(sPath)
    With CreateObject("Adodb.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile sPath
        .Position = 0
        GetBytes = .Read
        .Close
    End With
End Function