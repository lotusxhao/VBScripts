 Set olApp=CreateObject("Outlook.Application")
 Set olns=olApp.GetNameSpace("MAPI")
 Set objFolder=olns.GetDefaultFolder(10)
 
 For each item1 in objFolder.Items
         msgbox item1.subject    
 Next
 