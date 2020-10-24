# NotesFTP Script Library
© 2000 Paul D. Ray (pray@babson.edu)

This sample database contains a single script library, NotesFTP, that demonstrates how to implement basic FTP functionality in your applications. Using the methods and properties of the NotesFTPSession Class, you can perform the following actions in your client or server LotusScript modules against a remote FTP server (provided you have appropriate access):

- Navigate, enumerate, create, remove, and rename directories
- Upload, download, rename, and remove files. 

The NotesFTPSession Class makes calls to the Win32 Internet function library (wininet.dll), which is part of the Win32 family of operating systems. Consequently, this script library can only be used on the Win32 platform.* 


## Benefits of NotesFTP

If you've ever implemented FTP functionality in your applications using LotusScript, chances are it was done by making a call to the Shell function and running an FTP executable against a batch (.BAT) file containing FTP commands. Having passwords stored on the file system, even temporarily, can make an Administrator cringe, and "shelling" out to the command line is something many seasoned Developers do not want a part of.

NotesFTP provides a mechanism that overcomes both of these hurdles. For one, usernames and passwords can be kept in the confines of a Notes/Domino LotusScript module (e.g., agent), so they are secure. Secondly, the FTP calls are made "in process" rather than by launching a separate executable to do so, making the code more robust and efficient.

NotesFTP is also easy to use. Since it is implemented as a LotusScript Class, you are sheltered from making calls directly to the Win32 Internet API, and thus do not need to be concerned with the details surrounding calls to DLLs. Instead, the details are all encapsulated in an object that you manipulate.


## Implementing the NotesFTP Script Library

Before the methods and properties of the NotesFTP Script Library can be called from your application, it needs to be referenced in your LotusScript code module. There are two methods to accomplish this. The recommended option is to copy the script library to your application, and include the "Use" statement in the "Options" event of your LotusScript module:

```
Use "NotesFTP"
```

The other option is to export the script library to an .LSS file, comment out the "Option Public" line, and compile it into your application using the "%Include" statement in the "(Declarations)" event of your code module.

Accessing the methods and properties of the NotesFTPSession Class generally requires a (5)-step process that includes the following:

1.) Instantiate a NotesFTPSession object. This is done by assigning a variable of type NotesFTPSession with the "New" keyword.
2.) Connect to the FTP host. Use the Connect method to login to an FTP host.
3.) Perform any of the file and/or directory functions (e.g., GetFile method to download a file).
4.) Disconnect from the FTP host: Use the Disconnect method to logoff the FTP host. This is optional as "Delete" calls this method, but good practice. 
5.) Deference the NotesFTPSession object: Use the "Delete" keyword to deference an object. Do not simply let it fall out of scope.

The code snippet in *Figure 1*, which shows how to download a file called "readme.txt" from "ftp.lotus.com" (saving it as "readme2.txt"), exemplifies this process.

*Figure 1 - Downloading a file using the NotesFTPSession Class (Button script)*
```
Use "NotesFTP"

Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession

	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 .ChangeDirectory "pub/lotusweb/product/notesr5"
	 .GetFile "readme.txt", "c:\readme2.txt", FTP_TRANSFER_TYPE_ASCII
	 .Disconnect
	End With
	 
	Delete objFTP
	
End Sub 
```

The next section covers the methods and properties on the NotesFTPSession Class.


## NotesFTP API

Below are all the public methods and properties of the NotesFTP Script Library that can be called from a LotusScript module. Expand a section title for detailed information on a particular method or property:

Methods:

### NotesFTPSession.ChangeDirectory - Sets current directory on an FTP host.

### Syntax
NotesFTPSession.ChangeDirectory(dirName$)

### Elements
dirName$ - The name of the directory you wish to change to, relative to the current directory on the FTP host.

### Return Value
None. Raises a NOTESFTP_SETDIR_FAILED error on fail.

### Sample Usage
The following sample changes the current directory on the FTP host, and displays the directory name:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 .ChangeDirectory "pub"
	 Msgbox "Current Directory: " & .CurrentDirectory
	 .Disconnect
	End With
	
	Delete objFTP
	
End Sub
```

### NotesFTPSession.Connect - Logs in to an FTP host.

### Syntax
NotesFTPSession.Connect(serverName$, userName$, password$, flags&)

### Elements
serverName$ - The name of the FTP host to connect to (e.g., "ftp.lotus.com").
userName$ - The name of the user to log into the FTP host as (e.g., "anonymous").
password$ - The password of the user logging into the FTP host
flags& - Indicates whether you wish to use passive semantics. Set this to zero (0) if you are not sure. Otherwise, pass it the INTERNET_FLAG_PASSIVE constant.

### Return Value
None. Raises a NOTESFTP_CONNECT_FAILED error on fail.

### Sample Usage
This sample connects to "ftp.lotus.com" as an anonymous user:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 Msgbox "Connected? " & .IsConnected
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```



### NotesFTPSession.IsConnected - Indicates whether an application is connected to an FTP host.

### Syntax
NotesFTPSession.IsConnected

### Return Value
True on success, False on fail. No errors are raised from calling this property.

### Sample Usage
This sample returns the connection status (True or False) after logging into "ftp.lotus.com" as an anonymous user:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 Msgbox "Connected? " & .IsConnected
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```


### NotesFTPSession.CreateDirectory - Creates a new directory on an FTP host.

### Syntax
NotesFTPSession.CreateDirectory(dirName$)

### Elements
dirName$ - The name of the directory you wish to create, relative to the current directory on the FTP host. The user must have sufficient rights on the server to create the directory.

### Return Value
None. Raises a NOTESFTP_CREATEDIR_FAILED error on fail.

### Sample Usage
The following sample creates a new directory called "tempdir" on the "ftp.testdomain.com" host:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.testdomain.com", "jsmith", "abc", 0
	 .CreateDirectory "tempdir"
	 .Disconnect   
	End With
	
	Delete objFTP 
	
End Sub
```

### NotesFTPSession.DeleteFile - Deletes a file from an FTP host.

### Syntax
NotesFTPSession.DeleteFile(fileName$)

### Elements
fileName$ - The name of the file you wish to delete from the current directory on the FTP host. The user must have sufficient rights on the server to delete the file.

### Return Value
None. Raises a NOTESFTP_DELETEFILE_FAILED error on fail.

### Sample Usage
This sample deletes a file called "delete.me" from the FTP host's current directory:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession

	With objFTP
	 .Connect "ftp.testdomain.com", "jsmith", "password", 0
	 .DeleteFile "delete.me"
	 .Disconnect   
	End With

	Delete objFTP
	
End Sub
```

### NotesFTPSession.Dir - Retrieves a list of files and directory names from the current directory on an FTP host.

### Syntax
NotesFTPSession.Dir(dirSpec$)

### Elements
dirSpec$ - The file types to retrieve a list of (can include wildcards such as "*.*") from the current directory on the FTP host.

### Return Value
Variant containing an array of strings with filenames and directory names for the current directory on the FTP host. Use the UBound LotusScript function to find the number of names in the return value. This method does not raise any errors.

### Sample Usage
This sample displays a list of files in the "pub/lotusweb/product/notesr5" directory on the "ftp.lotus.com" host:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	Dim vFiles As Variant
	Dim nCt%
	Dim sMsg$, sDirPath$, sCRLF$
	
	Set objFTP=New NotesFTPSession
	sDirPath$="pub/lotusweb/product/notesr5"
	sCRLF$ = Chr(10) & Chr(13)
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 .ChangeDirectory sDirPath$
	 vFiles=.Dir("*.*")
	 .Disconnect  
	End With
	
	sMsg$="Directory listing of " & sDirPath$ & " (" & Cstr(Ubound(vFiles) + 1) & " files found):" & sCRLF$ & sCRLF$
	For nCt%=0 To Ubound(vFiles)
	 sMsg$=sMsg$ & vFiles(nCt%) & sCRLF$
	Next
	Msgbox sMsg$, 64, "NotesFTP Sample"
	
	Delete objFTP  
	
End Sub
```

### NotesFTPSession.Disconnect - Logs off an FTP host. Note: Delete calls this method as well upon deferencing a NotesFTPSession object.

### Syntax
NotesFTPSession.Disconnect

### Return Value
There are no return values and no errors raised by this method.

### Sample Usage
This sample returns the connection status (True or False) after logging off  "ftp.lotus.com":

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 .Disconnect
	 Msgbox "Connected? " & .IsConnected   
	End With
	
	Delete objFTP
	
End Sub
```


### NotesFTPSession.PutFile - Transfers a file to an FTP host.

### Syntax
NotesFTPSession.PutFile(localFile$, remoteFile$, transferType&)

### Elements
localFile$ - The name of the file on your local drive to upload. The user must have sufficient rights on the server to upload the file.
remoteFile$ - The name of the file to be saved as in the current directory of the FTP host.
transferType& - The flag indicating how the function will handle the upload. Pass FTP_TRANSFER_TYPE_BINARY for binary transfers, or FTP_TRANSFER_TYPE_ASCII for ASCII transfers.

### Return Value
None. Raises a NOTESFTP_PUTFILE_FAILED error on fail.

### Sample Usage
This sample uploads the local file "telnet.txt" to the FTP host's current directory using ASCII as the transfer type, saving it as "tn0529.txt" on the server:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.testdomain.com", "jsmith", "password", 0
	 .PutFile "c:\telnet.txt", "tn0529.txt", FTP_TRANSFER_TYPE_ASCII
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```

### NotesFTPSession.RemoveDirectory - Removes a directory from an FTP host.

### Syntax
NotesFTPSession.RemoveDirectory(dirName$)

### Elements
dirName$ - The name of the directory you wish to remove, relative to the current directory on the FTP host. The user must have sufficient rights on the server to delete the directory.

### Return Value
None. Raises a NOTESFTP_DELETEDIR_FAILED error on fail.

### Sample Usage
This sample removes a directory called "samples" from the current directory on the FTP host:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.mydomain.com", "jsmith", "lnotes", 0
	 .RemoveDirectory "samples"
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```

### NotesFTPSession.RenameFile - Renames a file on an FTP host.

### Syntax
NotesFTPSession.RenameFile(oldFileName$, newFileName$)

### Elements
oldFileName$ - The name of the file you wish to rename in the current directory on the FTP host. The user must have sufficient rights on the server to rename the file.
newFileName$ - The new name you would like to give the file.

### Return Value
None. Raises a NOTESFTP_RENAMEFILE_FAILED error on fail.

### Sample Usage
The following sample renames a file from NOTES.INI to NOTES.OLD in the current directory of the FTP host:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.mydomain.com", "jsmith", "lnotes", 0
	 .RenameFile "NOTES.INI", "NOTES.OLD"
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```

### Properties:

### NotesFTPSession.CurrentDirectory - Gets the current directory on an FTP host.

### Syntax
To Get: dirName$=NotesFTPSession.CurrentDirectory

### Return Value
String indicating the current directory on an FTP host. Raises a NOTESFTP_GETDIR_FAILED error on fail.

### Sample Usage
This sample returns the current directory on the "ftp.lotus.com" host:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 Msgbox "Current directory: " & .CurrentDirectory
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```

### NotesFTPSession.IsConnected - Indicates whether the application is connected to an FTP host.

### Syntax
NotesFTPSession.IsConnected

### Return Value
True on success, False on fail. No errors are raised from calling this property.

### Sample Usage
This sample returns the connection status (True or False) after logging into "ftp.lotus.com" as an anonymous user:

```
Sub Click(Source As Button)
	
	Dim objFTP As NotesFTPSession
	
	Set objFTP=New NotesFTPSession
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 Msgbox "Connected? " & .IsConnected
	 .Disconnect   
	End With
	
	Delete objFTP
	
End Sub
```


### Handling Errors

Some of the methods specified in the NotesFTP API raise runtime errors that you can trap in your code (see NotesFTP API for details). If you do not specify an "On Error" statement in your code, and a method of the NotesFTPSession object fails to execute properly, an error is raised by the object, and processing of the script stops. If you do trap errors using "On Error" though, you can shut down your references to the NotesFTPSession object more gracefully.

Figure 2 below shows a sample script that handles errors raised by a NotesFTPSession object.

Figure 2 - Handling runtime errors raised by a NotesFTPSession object.

```
Use "NotesFTP"

Sub Click(Source As Button)
	On Error Goto ErrHandler
	
	Dim objFTP As NotesFTPSession
	Dim vFiles As Variant
	Dim nCt%
	Dim sMsg$, sDirPath$, sCRLF$
	
	Set objFTP=New NotesFTPSession
	sDirPath$="pub/lotusweb/product/notesr5"
	sCRLF$ = Chr(10) & Chr(13)
	
	With objFTP
	 .Connect "ftp.lotus.com", "anonymous", "guest@testdomain.com", 0
	 .ChangeDirectory sDirPath$
	 vFiles=.Dir("*.*")
	 .Disconnect  
	End With
	
	sMsg$="Directory listing of " & sDirPath$ & ":" & sCRLF$ & sCRLF$
	
	For nCt%=0 To Ubound(vFiles)
	 sMsg$=sMsg$ & vFiles(nCt%) & sCRLF$
	Next
	Msgbox sMsg$, 64, "NotesFTP Sample"
	
Cleanup:
	If Not (objFTP Is Nothing) Then
	 Delete objFTP  
	End If
	
	Exit Sub
	
ErrHandler:
	Msgbox "[Error #" & Cstr(Err) & "]: " & Error$, 48, "NotesFTP Error"
	Resume Cleanup
	
End Sub
```

### Caveats and Limitations of This Sample

There are a few caveats and limitations of this sample that should be pointed out here. For one, as CERN proxies do not support FTP, applications that use a CERN proxy require additional development and are not supported in this sample. Second, this sample does not implement all FTP commands such as APPEND and TRACE. With additional work however, more functionality could be added to support these functions (refer to documentation of the FtpCommand Win32 Internet API and other related functions on MSDN Online at msdn.microsoft.com). Further, detailed file information such as size and date can be added with relatively little effort as well. As this is only a sample however, these pieces were left out in an effort to keep the scope small.


### Disclaimer

This sample is provided AS IS, with no warranties expressed or implied. The code herein is not supported in any way, shape, or form, and is not guaranteed to do anything. Therefore, the author cannot be held liable for any harm that may be caused by using this code. 


*Note: This code has only been tested on R5.0.3 clients and servers running Windows NT 4.0. It has not been used on R4.x clients or servers, Win2K, or on either of the "Wintendo" platforms (i.e., Win95 & Win98). 