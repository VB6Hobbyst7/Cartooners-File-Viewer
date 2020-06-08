'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Diagnostics.FileVersionInfo
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Reflection

'This modules contains this program's core procedures.
Public Module CoreModule
   Private Const PREFIX_SPACE_COUNT As Integer = &HB%   'Contains the number of spaces to remove from a music file's prefix.
   Private Const SUFFIX_SPACE_COUNT As Integer = &H1%   'Contains the number of spaces to remove from a music file's suffix.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim Data As List(Of Byte) = Nothing
         Dim Source As String = Nothing
         Dim Target As String = Nothing

         If GetCommandLineArgs.Count = 3 Then
            Source = GetCommandLineArgs(1)
            Target = GetCommandLineArgs(2)

            For Each FileO As FileInfo In My.Computer.FileSystem.GetDirectoryInfo(Source).GetFiles("*.mus")
               Data = New List(Of Byte)(File.ReadAllBytes(FileO.FullName))
               Console.Write(FileO.Name)

               If IsSherlock(Data) Then
                  Data.RemoveRange(&H0%, PREFIX_SPACE_COUNT)
                  Data.RemoveRange(Data.Count - SUFFIX_SPACE_COUNT, SUFFIX_SPACE_COUNT)
                  File.WriteAllBytes(Path.Combine(Target, FileO.Name), Data.ToArray())
                  Console.WriteLine(" - <CONVERTED>")
               Else
                  Console.WriteLine(" - <NOT CONVERTED>")
               End If
            Next FileO
         Else
            With My.Application.Info
               Console.WriteLine($"{ .Title} v{ .Version} - by: { .CompanyName}")
               Console.WriteLine()
               Console.WriteLine(.Description)
               Console.WriteLine()
               Console.WriteLine($"Usage ""{Path.GetFileName(GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileName)}"" SOURCE_DIRECTORY TARGET_DIRECTORY")
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine()
         Console.ForegroundColor = ConsoleColor.Red
         Console.WriteLine(ExceptionO.Message)
         Console.ResetColor()
         Console.ReadLine()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure checks whether the specified data indicates a Sherlock music file and returns the result.
   Private Function IsSherlock(Data As List(Of Byte)) As Boolean
      Try
         If Not Data.GetRange(&H0%, PREFIX_SPACE_COUNT).TrueForAll(Function(ByteV As Byte) ByteV = ToInt32(" "c)) Then Return False
         If Not Data.GetRange(Data.Count - SUFFIX_SPACE_COUNT, SUFFIX_SPACE_COUNT).TrueForAll(Function(ByteV As Byte) ByteV = ToInt32(" "c)) Then Return False
         Return True
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function
End Module
