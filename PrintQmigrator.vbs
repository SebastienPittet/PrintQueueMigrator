' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

'    Author: sebastien at pittet dot org
'    Date  : June 2015
'    Goal  : Migrate the network printer connections
'  Version : PrintQMigrator.vbs v.2015

' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

On Error Resume Next

'@@@@@@@@@@@@@
' MAIN PROGRAM
'@@@@@@@@@@@@@

Const TITLE = "PrintQMigrator v.1.1" 'Title for InputBoxes
Const ForReading = 1
Const DEFAULT_TEXTFILE = "ChangePrinter.txt"

Dim strDefaultPrinter	'store the name of the default printer
Dim InstalledPrinters	'Array of printer names
Dim strNoParams		'Text displayed if no parameters is given
Dim Textfile		'File which contains all printer information
Dim OldPrintQueues()	'Dynamic array to store old print queue names, from the text file
Dim NewPrintQueues()	'Dynamic array to store new print queue names, from the text file
Dim fso			'File System Object
Dim objTextFile		'Text file object
Dim strNextLine		'Line of the text file
Dim i			'Index used to loop...
Dim WshNetwork		'Use to work with the print queues (requested because no prnadmin.dll)

strNoParams = "This Script reads a text file and set " & _
              "migrate the print queues defined on this computer" & vbCrLf & vbCrLf & _
              "Type the path of the file containing the information the script needs." & vbCrLf & VbCrLf

'Get the command line args
  Set Parameters = Wscript.arguments

'If no command line arguments provided, prompt for file
  If Parameters.Count = 0 Then
    Textfile = InputBox(strNoParams,Title, GetThisFolderPath & "\" & DEFAULT_TEXTFILE)
  Else
    Textfile = Parameters.item(0)
  End If

  If Textfile = "" or Not Right(Textfile,4) = ".txt" or Not FileExist(Textfile) Then
     Error=MsgBox("No valid input file provided. Stopping the script now.",vbokonly, Title)
     WScript.Quit(1)
  End If

'Read the text file and import it in an Array
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile (TextFile, ForReading)

i=0
While not objTextFile.AtEndOfStream
	Redim Preserve OldPrintQueues(i)
	ReDim Preserve NewPrintQueues(i)
	strLine = objTextFile.Readline
	'Import only lines corresponding to Shared print queues.
	If Left(strLine,2) = "\\" Then
		OldPrintQueues(i) = Left(strLine,InStr(strline,";")-1)
		NewPrintQueues(i) = Mid(strline,InStr(strline,";")+1,Len(strline))
		i=i+1
	End If
Wend

objTextFile.Close 'Parameters file -> Closing

Set WshNetwork = CreateObject("WScript.Network")

'Store the name of the default Printer
strDefaultPrinter = DefaultPrinter

'Get all printer connections on this computer & user profile
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")
    
'Loop in printer collection of this workstation/user
For Each objPrinter in colInstalledPrinters
	If Left(objPrinter.Name, 2) = "\\" Then 'Work only On network printers
		'Search the corresponding printer and create it
		i = 0 'set the index at the beginning of the array (prepare to loop)
				
		Do Until i >= UBound(OldPrintQueues)
			If UCase(objPrinter.Name) = UCase(OldPrintQueues(i)) Then
				'Create the connection To the new printer if needed
				If Ucase(NewPrintQueues(i)) <> "" Then
					WshNetwork.AddWindowsPrinterConnection NewPrintQueues(i) 'Add the new print queue
					If UCase(objPrinter.Name) = UCase(strDefaultPrinter) Then 'Adapt the defaut print queue to the new one
						'Set the default Printer
						WshNetwork.SetDefaultPrinter NewPrintQueues(i)
					End If
				End if
				'Delete the old printer connection
				WshNetwork.RemovePrinterConnection OldPrintQueues(i)			
			End If
			i = i + 1
		Loop
	End If 'End of check for network printers
Next 'End of the loop through the printers of this user

Set WshNetwork = Nothing

'@@@@@@@@@@@
' Functions
'@@@@@@@@@@@

'--------------------------------------------------------------------

'Return the defaut printer
Function DefaultPrinter
	Dim strComputer
	Dim Result
	
	strComputer = "."
	Result = ""
	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colInstalledPrinters =  objWMIService.ExecQuery _
	    ("Select * from Win32_Printer")
	For Each objPrinter in colInstalledPrinters
		If objPrinter.Default = True Then
	    	Result = objPrinter.Name
		End If
	Next
	DefaultPrinter = Result
End Function

'--------------------------------------------------------------------
'Check If File Exist at a specified path (Boolean)
   Function FileExist (FileFullPath)
      Dim Fso
      Set Fso = CreateObject("Scripting.FileSystemObject")
      If (Fso.FileExists(FileFullPath)) Then
         FileExist = True
      Else
         FileExist = False
      End If
   End Function
'--------------------------------------------------------------------
'Get the path from where this script is executed.
   Function GetThisFolderPath()
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set file = fso.GetFile(wscript.scriptfullname)
      GetThisFolderPath=File.ParentFolder
   End Function
'--------------------------------------------------------------------

