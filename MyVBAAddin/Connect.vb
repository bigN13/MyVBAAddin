Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Extensibility

<ComVisible(True), Guid("C3BC0970-0D67-4518-86C3-E3D693E55287"), ProgId("MyVBAAddin.Connect")> _
Public Class Connect
   Implements Extensibility.IDTExtensibility2

   'Private _VBE As VBAExtensibility.VBE
   'Private _AddIn As VBAExtensibility.AddIn

   Private Sub OnConnection(Application As Object, ConnectMode As Extensibility.ext_ConnectMode, _
      AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

      Try

         '_VBE = DirectCast(Application, VBAExtensibility.VBE)
         '_AddIn = DirectCast(AddInInst, VBAExtensibility.AddIn)

         Select Case ConnectMode

            Case Extensibility.ext_ConnectMode.ext_cm_Startup
               ' OnStartupComplete will be called

            Case Extensibility.ext_ConnectMode.ext_cm_AfterStartup
               InitializeAddIn()

         End Select

      Catch ex As Exception

         MessageBox.Show(ex.ToString())

      End Try

   End Sub

   Private Sub OnDisconnection(RemoveMode As Extensibility.ext_DisconnectMode, _
      ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

   End Sub

   Private Sub OnStartupComplete(ByRef custom As System.Array) _
      Implements IDTExtensibility2.OnStartupComplete

      InitializeAddIn()

   End Sub

   Private Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

   End Sub

   Private Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

   End Sub

   Private Sub InitializeAddIn()

      'MessageBox.Show(_AddIn.ProgId & " loaded in VBA editor version " & _VBE.Version)
      MessageBox.Show("My Addin is loaded in VBA editor version ")

   End Sub

End Class