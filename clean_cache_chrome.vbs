Dim str
Dim path
Dim SizeCache
UserProfilePath = "C:\users"
CachePath = "\AppData\Local\Google\Chrome\User Data\Default\Cache"

Set FSO = CreateObject("Scripting.FileSystemObject")

Set objFolderPofile = FSO.GetFolder(UserProfilePath)

For Each SubFolder In objFolderPofile.SubFolders
	 str = CStr(SubFolder)
	 path = FSO.BuildPath(str, CachePath)
	 
	 IF FSO.FolderExists(path) = -1   Then
		                                             WScript.Echo path 'debug
		Set CacheFolder = FSO.GetFolder(path)
		Set colFiles = CacheFolder.Files
		For Each objFile in colFiles
			SizeCache = SizeCache + objFile.Size		
			objFile.Delete(false)                               'Delete Files
		Next
	End If
Next
