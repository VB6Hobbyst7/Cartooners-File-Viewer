'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'This structure defines the executable to be edited.
   Public Structure ExecutableStr
      Public Data As List(Of Byte)   'Defines the executable's data.
      Public Path As String          'Defines the executable's path.
   End Structure

   Public WithEvents Disassembler As New DisassemblerClass   'Contains a reference to the disassembler.

   Public Const CHARACTERS_PER_BYTE As Integer = &H3%   'Defines the number of hexadecimal digits including the number of spaces per byte.

   'This procedure manages data for the executable being edited.
   Public Function Executable(Optional ExecutablePath As String = Nothing) As ExecutableStr
      Try
         Static CurrentExecutable As ExecutableStr = Nothing

         If Not ExecutablePath = Nothing Then
            CurrentExecutable = New ExecutableStr With {.Data = New List(Of Byte)(File.ReadAllBytes(ExecutablePath)), .Path = ExecutablePath}
         End If

         Return CurrentExecutable
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified bytes as hexadecimal text.
   Public Function GetHexadecimals(Data As List(Of Byte)) As String
      Try
         Dim Hexadecimals As String = Nothing
         Dim Offset As New Integer
         Dim Length As New Integer

         If Data IsNot Nothing Then
            Hexadecimals = String.Join(" "c, From ByteO As Byte In Data Select $"{ByteO:X2}")
            If Regions()?.Count > 0 Then
               For Each Region As RegionClass In Regions()
                  Offset = Region.Offset * CHARACTERS_PER_BYTE
                  Length = Region.Length * CHARACTERS_PER_BYTE
                  Hexadecimals = $"{Hexadecimals.Substring(0, Offset)}{String.Join(" "c, Enumerable.Repeat("##", Length \ CHARACTERS_PER_BYTE))} {Hexadecimals.Substring(Offset + Length)}"
               Next Region
            End If
         End If

         Return Hexadecimals
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception) Handles Disassembler.HandleError
      Try
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      Catch
         Application.Exit()
      End Try
   End Sub

   'This procedure returns information about this program.
   Public Function ProgramInformation() As String
      Try
         With My.Application.Info
            Return $"{ .Title} v{ .Version} - by { .CompanyName}"
               End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
