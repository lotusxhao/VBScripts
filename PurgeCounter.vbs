Const ForReading = 1
Const ForWriting = 2

Set objFSO1 = CreateObject("Scripting.FileSystemObject")
Set objFile1 = objFSO1.OpenTextFile("C:\Windows Scripts\OUTPUT\PurgeCount.csv", ForWriting, TRUE)

DATUM = 366
PARENTLOOP_TERM = 25380
CHILDLOOP_TERM = 1000

loopRemainder = PARENTLOOP_TERM Mod CHILDLOOP_TERM
parentLoopTerm = PARENTLOOP_TERM - loopRemainder

For i = 1 to 25
    For j = 1 to 1000
        objFile1.Write DATUM & "," & vbcrlf
    Next
	
	DATUM = DATUM + 1
Next

For i = 1 to 380
    objFile1.Write DATUM & "," & vbcrlf
Next

objFile1.Close 

