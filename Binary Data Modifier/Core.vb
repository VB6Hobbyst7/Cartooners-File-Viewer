'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.Convert
Imports System.Diagnostics
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module CoreModule

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim Arguments As String = Nothing
         Dim Data As List(Of Byte) = Nothing
         Dim InputPath As String = Nothing
         Dim Length As Integer = Nothing
         Dim ModificationType As Char = Nothing
         Dim NewData As List(Of Byte) = Nothing
         Dim Offset As Integer = Nothing
         Dim OpenWith As String = Nothing
         Dim OutputPath As String = Nothing
         Dim ScriptPath As String = If(GetCommandLineArgs.Count > 1, GetCommandLineArgs.Last(), Nothing)

         If ScriptPath = Nothing Then
            Console.WriteLine($"Specify a script file as a command line argument.")
            Console.ReadLine()
         Else
            Using ScriptFile As New StreamReader(ScriptPath)
               With ScriptFile
                  InputPath = .ReadLine()
                  OutputPath = .ReadLine()
                  Offset = Integer.Parse(.ReadLine())
                  Length = Integer.Parse(.ReadLine())
                  ModificationType = .ReadLine().Trim().ToUpper().ToCharArray().First()
                  If ModificationType = "R"c Then NewData = New List(Of Byte)(From [Byte] In .ReadLine().Trim().Split(" "c) Select ToByte([Byte], fromBase:=16))
                  If Not .EndOfStream() Then OpenWith = .ReadLine()
                  If Not .EndOfStream() Then Arguments = .ReadLine()
               End With
            End Using

            Data = New List(Of Byte)(File.ReadAllBytes(InputPath))

            Select Case ModificationType
               Case "I"c
                  Data = InvertByteBits(Data, Offset, Length)
               Case "R"c
                  Data = ReplaceBytes(Data, NewData, Offset, Length)
               Case "Z"c
                  Data = ZeroBytes(Data, Offset, Length)
               Case Else
                  Console.WriteLine($"Unknown modification type: ""{ModificationType}"".")
            End Select

            File.WriteAllBytes(OutputPath, Data.ToArray())

            If OpenWith IsNot Nothing Then Process.Start(New ProcessStartInfo With {.FileName = OpenWith, .Arguments = Arguments, .WindowStyle = ProcessWindowStyle.Normal})
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine($"Error: {ExceptionO.Message}")
         Console.ReadLine()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure the bits of the specified number of bytes at the specified offset inside the specified data and returns the result.
   Private Function InvertByteBits(Data As List(Of Byte), Offset As Integer, Length As Integer) As List(Of Byte)
      Try
         Dim Position As Integer = Offset

         Do Until Position >= Offset + Length OrElse Position >= Data.Count
            Data(Position) = ToByte(Data(Position) Xor &HFF%)
            Position += &H1%
         Loop

         Return Data
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure replaces the data of the specified length at the specified offset inside the data specified with the new data specified and returns the result.
   Private Function ReplaceBytes(Data As List(Of Byte), NewData As List(Of Byte), Offset As Integer, Length As Integer) As List(Of Byte)
      Try
         Dim Position As Integer = Offset

         Do Until Position + NewData.Count >= Offset + Length OrElse Position >= Data.Count
            For SubPosition As Integer = Position To Position + (NewData.Count - &H1%)
               If SubPosition >= Data.Count Then Exit For
               Data(SubPosition) = NewData(SubPosition - Position)
            Next SubPosition
            Position += NewData.Count
         Loop

         If Position < Offset + (Length - &H1%) AndAlso Position < Data.Count Then
            For SubPosition As Integer = Position To Offset + (Length - &H1%)
               If SubPosition >= Data.Count Then Exit For
               Data(SubPosition) = NewData(SubPosition - Position)
            Next SubPosition
         End If

         Return Data
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure zeroes the specified number of bytes at the specified offset inside the specified data and returns the result.
   Private Function ZeroBytes(Data As List(Of Byte), Offset As Integer, Length As Integer) As List(Of Byte)
      Try
         If Offset + Length >= Data.Count Then Length = Data.Count - Offset

         Data.RemoveRange(Offset, Length)
         Data.InsertRange(Offset, Enumerable.Repeat(ToByte(&H0%), Length))

         Return Data
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function
End Module
