Dim ProjObj
Dim ProfileMissing
Dim ProfileMissingDoublePwa
Dim i
Dim memOfArray() 
Dim arrayPwa2() 
Dim arrayPwa5() 
i = 0
k1 = 0
k2 = 0
sizeOfArrayPwa2 = 0
sizeOfArrayPwa5 = 0
Const c_ServerPwa5 = "http://project/pwa5"
Const c_ServerPwa2 = "http://project/pwa2"
Const DefaultProfile = "Компьютер"
ProfileMissing = True
ProfileMissingDoublePwa = True
ServerPwa2 = "pwa2"
ServerPwa5 = "pwa5"
ouProject = "OU=Project Server,OU=Группы доступа,OU=Сервис,DC=sminex,DC=com"
'---------------<<<  End advertisement variables >>>----------------------------


Set objSysInfo = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
objMembetOf = objUser.Get("memberOf")

'-----------<<<<<  Algorithm for re-count and read Notes >>>>>>----------------------
For Each InGroup In objMembetOf
	strN = InStr(1, InGroup, ouProject)		
	If strN > 0 Then
		Set objGrp = GetObject("LDAP://" & InGroup)				                                            		
		ReDim Preserve memOfArray(k1) 
		memOfArray(k1) = objGrp.Get("info") 
		                                                                                        'WScript.Echo " " & memOfArray(k1)       ' For Debug		
		If memOfArray(k1) = ServerPwa2 Then
	    ReDim Preserve arrayPwa2(k1)
		arrayPwa2(k1) = memOfArray(k1)
		sizeOfArrayPwa2 = sizeOfArrayPwa2 + 1		                                             'WScript.Echo " " & arrayPwa2(k1)        ' For Debug
		k1 = k1 + 1			
	Else  		
	    ReDim Preserve arrayPwa5(k2)
		arrayPwa5(k2) = memOfArray(k1)
		sizeOfArrayPwa5 = sizeOfArrayPwa5 + 1		                                             'WScript.Echo " " & arrayPwa5(k2)        ' For Debug
		k2 = k2 + 1
	    k1 = k1 + 1
		
	End If	
			
	End If
		
		i = i + 1
		Set objGrp = Nothing
Next


 '-----------<<<<<  Algorithm for pwa2 and pwa5  >>>>>>----------------------
 If sizeOfArrayPwa2 > 0 And sizeOfArrayPwa5 > 0 Then
  Set ProjObj = CreateObject("msproject.application")  
  For i = 1 To ProjObj.Profiles.Count 
    If ProjObj.Profiles.Item(i).Server = c_ServerPwa5 Then
	ProfileMissingDoublePwa = False
    End If
	Next
	
	 If ProfileMissingDoublePwa = True Then
     ProjObj.Profiles.Add ServerPwa5, c_ServerPwa5
     End If
	
  For i = 1 To ProjObj.Profiles.Count 
    If ProjObj.Profiles.Item(i).Server = c_ServerPwa2 Then
	ProfileMissingDoublePwa = False
    End If
	Next
	
	If ProfileMissingDoublePwa = True Then
    ProjObj.Profiles.Add ServerPwa2, c_ServerPwa2
     End If
	 
	 For i = 1 To ProjObj.Profiles.Count
    If ProjObj.Profiles.Item(i).Name = DefaultProfile Then
        ProjObj.Profiles.DefaultProfile = ProjObj.Profiles.Item(i)
    End If
     Next
	 ProjObj.Quit 
     Set ProjObj=Nothing  
                                                                                 'WScript.Echo "Default Computer"   'For Debug
   
   '-------<<<<<  Algorithm for pwa2 >>>>>>----------------------
   Else if sizeOfArrayPwa2 > 0 Then
   Set ProjObj = CreateObject("msproject.application")
        For i = 1 To ProjObj.Profiles.Count
       If ProjObj.Profiles.Item(i).Server = c_ServerPwa2 Then 
	     ProfileMissing = False
		 End If
        Next
		
	 If ProfileMissing = True Then
    ProjObj.Profiles.Add ServerPwa2, c_ServerPwa2
     End If
	 
	 For i = 1 To ProjObj.Profiles.Count
    If ProjObj.Profiles.Item(i).Name = ServerPwa2 Then
        ProjObj.Profiles.DefaultProfile = ProjObj.Profiles.Item(i)
    End If
     Next
	 ProjObj.Quit 
     Set ProjObj=Nothing
   
                                                                                  'WScript.Echo "pwa2"   'For Debug
	  
    '-------<<<<<  Algorithm for pwa5  >>>>>>----------------------
   Else if sizeOfArrayPwa5 > 0 Then 
   Set ProjObj = CreateObject("msproject.application")   
     For i = 1 To ProjObj.Profiles.Count
       If ProjObj.Profiles.Item(i).Server = c_ServerPwa5 Then 
	     ProfileMissing = False
		 End If
     Next
	 
	 If ProfileMissing = True Then
    ProjObj.Profiles.Add ServerPwa5, c_ServerPwa5
     End If
	 
	 For i = 1 To ProjObj.Profiles.Count
    If ProjObj.Profiles.Item(i).Name = ServerPwa5 Then
        ProjObj.Profiles.DefaultProfile = ProjObj.Profiles.Item(i)
    End If
	  next
	  ProjObj.Quit 
      Set ProjObj=Nothing
	                                                                              'WScript.Echo "pwa5"    'For Debug
    End If
	
   
   End If
   End If
  