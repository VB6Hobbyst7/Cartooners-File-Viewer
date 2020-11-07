'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.Linq
Imports System.Windows.Forms

'This module contains this program's interface.
Public Class InterfaceWindow
   'This procedure initializes this window.
   Public Sub New()
      Try
         InitializeComponent()

         My.Application.ChangeCulture("en-US")

         With My.Computer.Screen.WorkingArea
            Me.Size = New System.Drawing.Size(CInt(.Width / 1.1), CInt(.Height / 1.1))
         End With

         Me.Text = ProgramInformation()

         SplitterBox.SplitterDistance = CInt(SplitterBox.Height / 2)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to load the file dropped into the data box.
   Private Sub DataBox_DragDrop(sender As Object, e As DragEventArgs) Handles DataBox.DragDrop
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then LoadExecutable(DirectCast(e.Data.GetData(DataFormats.FileDrop), String()).First)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles objects being dragged into the data box.
   Private Sub DataBox_DragEnter(sender As Object, e As DragEventArgs) Handles DataBox.DragEnter
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then e.Effect = DragDropEffects.All
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to disassemble the code at the newly selected position in the executable's data.
   Private Sub DataBox_SelectionChanged(sender As Object, e As EventArgs) Handles DataBox.Click, DataBox.KeyUp
      Try
         Dim DisassemblyCode As New List(Of String)
         Dim Length As Integer = Nothing
         Dim Position As Integer = &H0%
         Dim SelectedExecutableCode As New List(Of Byte)

         DataBox.Select(AlignToBytes(DataBox.SelectionStart - CHARACTERS_PER_BYTE), AlignToBytes(DataBox.SelectionLength))

         Length = DataBox.SelectionLength \ CHARACTERS_PER_BYTE
         SelectedExecutableCode = Executable().Data?.GetRange(DataBox.SelectionStart \ CHARACTERS_PER_BYTE, Length)
         Do Until Position >= Length OrElse My.Application.OpenForms.Count = 0
            DisassemblyCode.Add(Disassembler.Disassemble(SelectedExecutableCode, Position))
            Application.DoEvents()
         Loop

         DisassemblyBox.Text = String.Join(NewLine, DisassemblyCode)
         UpdateStatusBar()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure closes this program when this window is closed.
   Private Sub InterfaceWindow_Closing(sender As Object, e As EventArgs) Handles Me.Closing
      Try
         Application.Exit()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure further initializes this window.
   Private Sub InterfaceWindow_Load(sender As Object, e As EventArgs) Handles Me.Load
      Try
         If GetCommandLineArgs.Count > 1 Then LoadExecutable(GetCommandLineArgs.Last())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure adjusts this window's controls to its new size.
   Private Sub InterfaceWindow_Resize(sender As Object, e As EventArgs) Handles Me.Resize
      Try
         SplitterBox.Height = Me.ClientSize.Height - StatusBar.Height
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This program displays a dialog requesting the user to select an executable to be edited.
   Private Sub LoadExecutableMenu_Click(sender As Object, e As EventArgs) Handles LoadExecutableMenu.Click
      Try
         With New OpenFileDialog
            If .ShowDialog() = DialogResult.OK Then LoadExecutable(ExecutablePath:= .FileName)
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to load the specified regions.
   Private Sub LoadRegionsMenu_Click(sender As Object, e As EventArgs) Handles LoadRegionsMenu.Click
      Dim Overlaps As String = Nothing
      With New OpenFileDialog
         If .ShowDialog() = DialogResult.OK Then
            Regions(RegionPath:= .FileName)
            Overlaps = CheckForRegionOverlap()
            If Not Overlaps = Nothing Then MessageBox.Show(Overlaps, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
            DataBox.Text = GetHexadecimals(Executable.Data)
            DataBox.Select(0, 0)
         End If
      End With
   End Sub

   'This procedure gives the command to close this window.
   Private Sub QuitMenu_Click(sender As Object, e As EventArgs) Handles QuitMenu.Click
      Try
         Me.Close()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure aligns the specified position/length inside the executable's hexadecimal data representation with its raw byte data.
   Private Function AlignToBytes(Value As Integer) As Integer
      Try
         Return ((Value \ CHARACTERS_PER_BYTE) + &H1%) * CHARACTERS_PER_BYTE
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure gives the command to load the specified executable.
   Private Sub LoadExecutable(ExecutablePath As String)
      Try
         Executable(ExecutablePath:=ExecutablePath)
         DataBox.Text = GetHexadecimals(Executable.Data)
         DataBox.Select(0, 0)
         Regions(, Reset:=True)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure updates the status bar.
   Private Sub UpdateStatusBar()
      Try
         Dim Position As Integer = DataBox.SelectionStart \ CHARACTERS_PER_BYTE
         Dim RegionList As List(Of RegionClass) = Nothing

         StatusPositionLabel.Text = $"Position: {Position }/{Executable.Data?.Count - &H1%}"
         If Regions()?.Count > 0 Then
            RegionList = New List(Of RegionClass)(From Region In Regions() Where Position >= Region.Offset AndAlso Position <= Region.EndO)

            If RegionList.Count = 0 Then
               StatusRegionLabel.Text = "Region:"
            Else
               StatusRegionLabel.Text = $"Region: {RegionList.First.Description}"
            End If
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Class
