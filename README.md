# HOWTO: Create an add-in for the VBA editor (32-bit or 64-bit) of Office with Visual Studio .NET.

https://youtu.be/y81Aq4bebZU

Traditionally add-ins for the 32-bit VBA editor of Microsoft Office were created with Visual Basic 6.0, which can generate 32-bit COM (ActiveX) DLLs. However, the VBA editor of the 64-bit version of Microsoft Office 2010 only supports 64-bit COM add-ins, so it is not possible to create add-ins for that version using Visual Basic 6.0.

This article explains how to create an add-in for the 64-bit VBA editor of Office (and for the 32-bit VBA editor) using Visual Studio .NET and the .NET Framework 2.0, since .NET dlls can be 32-bit or 64-bit and both can be registered for COM-Interop.
