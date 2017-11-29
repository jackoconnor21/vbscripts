'==========================================================================================================================='
'																															'
'	 Author: Jack O'Connor																									'
'  Description: Recursively search directory X for file type Y	and write to an external file								'
'																															'
'==========================================================================================================================='

' The directory to begin the search'
Const directoryPath = "C:\"
Const fileType = "exe"
Const fileName = "C:\filesearchinfo.txt"

' Create file system object and get the folder from that object based on the given path'
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set searchLocation = objFSO.GetFolder(directoryPath)
Set myFile = objFSO.OpenTextFile(fileName, 2, True)

' Start a count at 0'
Count = 0

' Add in some useful information at the top of the text file'
myFile.Write(_
	"File type searched for: ." & fileType & vbCrLf & _ 
	"Directory searched: " & directoryPath & vbCrLf & _ 
	"------------------------------------------------" & vbCrLf _
)

' Call search folder function for the first time and pass the folder to search'
searchFolder(searchLocation)

' Close the stream to the textfile'
myFile.Close

' Open a stream to READ the text file and read all content'
Set myFile = objFSO.OpenTextFile(fileName, 1)
myFileContent = myFile.ReadAll

' Open a stream to write to the text file and prepend the number of files counted'
Set myFile = objFSO.OpenTextFile(fileName, 2, True)
myFile.WriteLine("Number of files found: "&Count)
myFile.Write(myFileContent)
myFile.Close

Sub searchFolder(folder)
	' If the folder is not accessible then do not run a search (as this wont work)'
	If IsAccessible(folder) Then
		' Search all files within the folder'
		For Each oFile In folder.Files
			' If the extension name is exe then we will write out the file name and path'
			If objFSO.getextensionname(oFile.path) = fileType Then 
				Count = Count + 1
				myFile.Write(vbCrLf & "File #" & Count & vbCrLf & "Name: " & oFile.Name & vbCrLf & "Path: " & oFile.Path & vbCrLf)
			End If

	    Next 

	    ' Loop over each sub folder of this current folder'
	    For Each objFolder In folder.SubFolders
	    	' Re-enter the current function, this time passing the sub folder instead'
		    searchFolder objFolder
		Next
	End If
End Sub

' Function for checking the number of subfolders
Function IsAccessible(oFolder)
	On Error Resume Next
	IsAccessible = oFolder.SubFolders.Count >= 0
End Function
