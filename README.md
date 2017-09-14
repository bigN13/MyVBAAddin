# HOWTO: Create an add-in for the VBA editor (32-bit or 64-bit) of Office with Visual Studio .NET.

https://youtu.be/y81Aq4bebZU

Traditionally add-ins for the 32-bit VBA editor of Microsoft Office were created with Visual Basic 6.0, which can generate 32-bit COM (ActiveX) DLLs. However, the VBA editor of the 64-bit version of Microsoft Office 2010 only supports 64-bit COM add-ins, so it is not possible to create add-ins for that version using Visual Basic 6.0.

This article explains how to create an add-in for the 64-bit VBA editor of Office (and for the 32-bit VBA editor) using Visual Studio .NET and the .NET Framework 2.0, since .NET dlls can be 32-bit or 64-bit and both can be registered for COM-Interop.

The COM type libraries used by add-ins of the VBA editor of Microsoft Office are the following:

Microsoft Add-In Designer ("MSADDNDR.DLL")
OLE Automation ("stdole2.tlb")
Microsoft Office <version> Object Library ("mso.dll")
Microsoft Visual Basic for Applications Extensibility 5.3 ("vbe6ext.olb")
  
To create an add-in for the 64-bit or 32-bit VBA editor of Microsoft Office follow these steps:

Using Visual Studio 2005 or higher create a VB.NET Class Library project named "MyVBAAddin" using .NET Framework 2.0. It is possible to use higher versions of the .NET Framework, but they are less likely to be installed on the machine of the end users so the setup of the add-in should install it. It is possible to use C# too.
Generate the required Interop assemblies as explained in the article HOWTO: Generate Interop assemblies to create an add-in for the VBA editor (32-bit or 64-bit) of Office with Visual Studio .NET.
Add those Interop assemblies as references to the project, using the "Project", "Add Reference..." menu, "Browse" tab. When the project is built, these Interop assemblies will be copied to the output folder (bin\Debug) along with the add-in assembly. All of them should be installed by the setup in the destination folder with the add-in assembly.
In the Project properties window of the project:
In the "Application" tab, ensure that both the Assembly name and Root namespace are set to "MyVBAAddin".
In the "Compile" tab, ensure that the "Register for COM interop" checkbox is not checked. The assembly will be registered for COM Interop manually later with the proper regasm.exe tool. This checkbox would only register the add-in dll as 32-bit COM library, not as 64-bit COM library.
In the "Compile" tab, "Advanced Compile Options" button, ensure that the "Target CPU" combobox is set to "AnyCPU", which means that the assembly can be executed as 64-bit or 32-bit, depending on the executing .NET Framework that loads it.
In the "Signing" tab, check the "Sign the assembly" checkbox and choose the strong name key file that you used to generate the Interop assemblies.
Rename the Class1.vb class to Connect.vb.
Paste in the Connect.vb file the code from the respository

Build the assembly.
To register the .NET assembly as COM component open a DOS (Command) window with admin rights (locate the C:\Windows\System32\cmd.exe file and right-click its "Run as administrator" context menu) and type:
To register the .NET assembly as 64-bit component:

C:\Windows\Microsoft.NET\Framework64\v2.0.50727\regasm.exe /codebase "<path-to-assembly>\MyVBAAddin.dll"
 
To register the .NET assembly as 32-bit component:

C:\Windows\Microsoft.NET\Framework\v2.0.50727\regasm.exe /codebase "<path-to-assembly>\MyVBAAddin.dll"

You should get a "Types registered successfully" output message.
To register the .NET assembly as VBA add-in:
To register it as add-in for the VBA editor 64-bit create the registry key:

HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\MyVBAAddIn.Connect

To register it as add-in for the VBA editor 32-bit create the registry key:

HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\MyVBAAddIn.Connect

Inside that registry key, create the following Names and Values:

Name	Type	Value
FriendlyName	REG_SZ	My VBA Add-in
Description	REG_SZ	My VBA Add-in
LoadBehavior	DWORD 32-bit	0

Notice that VBA add-ins are registered only for the current user; they can not be registered in the HKEY_LOCAL_MACHINE hive for all users.

Open a Office application.
Open its VBA editor.
Go to the "Add-Ins", "Add-In Manager..." menu to check that the add-in is correctly registered.
Load the add-in. You should see the message box "MyVBAAddin.Connect loaded in VBA editor version m.n"

Aknowledgements

https://www.mztools.com/articles/2012/MZ2012011.aspx

https://www.mztools.com/articles/2012/MZ2012013.aspx
