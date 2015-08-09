# PrintQmigrator for Windows

### General Description
Long time ago, during the old Windows NT4 time, Microsoft provided ChangePrint.exe as a part of the Windows NT4 resource kit. This tool could be used in the login script and was based on a text file which contained all "Old/New print queues definitions". **The aim was to manage the connected print queues on the user profiles (update things for the user, remove old printer connections, do some housekeeping)**.

However, ChangePrint does not run anymore from Windows XP and newer Windows OS and Microsoft provides no solution to this problem. Their official answer is: Windows XP support VB scripting, you only need to write some code.

That’s exactly what I did, trying to translate the ChangePrint in VBScript. I tested my script when I migrated a big print server and it did the job! I called it PrintQmigrator and decided to share it with the community.

### Usage
You can call PrintQmigrator from a login script or deploy it with a GPO. You can also put the text file containing all Old/New print queues definitions in the same location or on a file share (must be accessible in READ for each user). You can also have as many text files as you want!

### Synthax information
```
PrintQmigrator.vbs <TextFileFullPath.txt>
```

Exemple:
```
PrintQmigrator.vbs \\FileServer\Share\Printers.txt
```

For testing purpose, you can also double click on the file PrintQmigrator.vbs and it will ask you for a text file where the print queues are defined.

### Structure of parameters file
The PrintQmigrator text file is a basic CSV file, where the separator MUST be a semi-colon (;).

Below, you will find a good sample of a text file, that shows you can migrate the print queue on another server keeping the print queues' names or not. You can also keep the server but just rename the queues:

```
\\OldServer\OldPrintQueue;\\NewServer\NewPrintQueue
\\OldServer\OldPrintQueue1;\\OldServer\NewPrintQueue1
\\OldServer2\OldPrintQueue2;\\NewServer2\NewPrintqueue2
```

PrintQmigrator will only read the lines beginning with double backslashes. So you can easily comment the file with lines above or below the print queues definition.

### Print queue removal
If you only need to disconnect a print queue and __NOT__ to replace with another, just let the "new print queue" empty, i.e :

```
\\oldserver\oldprintqueue;
```

