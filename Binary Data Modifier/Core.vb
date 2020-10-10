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
         Dim ModifiedData As List(Of Byte) = Nothing
         Dim ModificationType As Char = Nothing
         Dim NewData As List(Of Byte) = Nothing
         Dim Offset As Integer = Nothing
         Dim OpenWith As String = Nothing
         Dim OutputPath As String = Nothing
         Dim ScriptPath As String = If(GetCommandLineArgs.Count > 1, GetCommandLineArgs.Last(), Nothing)
         Dim Substitutions As Dictionary(Of Byte, Byte) = Nothing

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

                  Select Case ModificationType
                     Case "R"c
                        NewData = New List(Of Byte)(From [Byte] In .ReadLine().Trim().Split(" "c) Select ToByte([Byte], fromBase:=16))
                     Case "S"c
                        Substitutions = New Dictionary(Of Byte, Byte)
                        For Each Pair As String In From BytePair As String In .ReadLine().Trim().Split(" "c)
                           With Pair.Split("="c)
                              Substitutions.Add(ToByte(.First, fromBase:=16), ToByte(.Last, fromBase:=16))
                           End With
                        Next Pair
                  End Select

                  If Not .EndOfStream() Then OpenWith = .ReadLine()
                  If Not .EndOfStream() Then Arguments = .ReadLine()
               End With
            End Using

            Data = New List(Of Byte)(File.ReadAllBytes(InputPath))
            ModifiedData = New List(Of Byte)(Data)

            Select Case ModificationType
               Case "I"c
                  InvertByteBits(ModifiedData, Offset, Length)
               Case "R"c
                  ReplaceBytes(ModifiedData, Offset, Length, NewData)
               Case "S"c
                  SubstituteBytes(ModifiedData, Offset, Length, Substitutions)
               Case "Z"c
                  ZeroBytes(ModifiedData, Offset, Length)
               Case Else
                  ModifiedData = Nothing
                  Console.WriteLine($"Unknown modification type: ""{ModificationType}"".")
                  Console.ReadLine()
            End Select

            If ModifiedData IsNot Nothing Then
               If ModifiedData.SequenceEqual(Data) Then
                  Console.WriteLine("No changes were made.")
                  Console.ReadLine()
               Else
                  File.WriteAllBytes(OutputPath, ModifiedData.ToArray())
                  If OpenWith IsNot Nothing Then Process.Start(New ProcessStartInfo With {.FileName = OpenWith, .Arguments = Arguments, .WindowStyle = ProcessWindowStyle.Normal})
               End If
            End If
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

   'This procedure inverts the bits of the specified number of bytes at the specified offset inside the specified data.
   Private Sub InvertByteBits(Data As List(Of Byte), Offset As Integer, Length As Integer)
      Try
         Dim Position As Integer = Offset

         Do Until Position >= Offset + Length OrElse Position >= Data.Count
            Data(Position) = ToByte(Data(Position) Xor &HFF%)
            Position += &H1%
         Loop
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure replaces the data of the specified length at the specified offset inside the data specified with the new data specified.
   Private Sub ReplaceBytes(Data As List(Of Byte), Offset As Integer, Length As Integer, NewData As List(Of Byte))
      Try
         Dim Position As Integer = Offset

         Do Until Position + NewData.Count >= Offset + Length OrElse Position >= Data.Count
            For SubPosition As Integer = Position To Position + (NewData.Count - &H1%)
               If SubPosition >= Data.Count Then Exit For
               Data(SubPosition) = NewData(SubPosition - Position)
            Next SubPosition
            Position += NewData.Count
         Loop

         If Position < Offset + Length AndAlso Position < Data.Count Then
            For SubPosition As Integer = Position To Offset + (Length - &H1%)
               If SubPosition >= Data.Count Then Exit For
               Data(SubPosition) = NewData(SubPosition - Position)
            Next SubPosition
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure substitutes the specified number of bytes at the specified offset inside the specified data using the specified dictionary.
   Private Sub SubstituteBytes(Data As List(Of Byte), Offset As Integer, Length As Integer, Substitutions As Dictionary(Of Byte, Byte))
      Try
         Dim Position As Integer = Offset

         Do Until Position >= Offset + Length OrElse Position >= Data.Count
            If Substitutions.ContainsKey(Data(Position)) Then Data(Position) = Substitutions(Data(Position))
            Position += &H1%
         Loop
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure zeroes the specified number of bytes at the specified offset inside the specified data.
   Private Sub ZeroBytes(Data As List(Of Byte), Offset As Integer, Length As Integer)
      Try
         If Offset + Length >= Data.Count Then Length = Data.Count - Offset

         Data.RemoveRange(Offset, Length)
         Data.InsertRange(Offset, Enumerable.Repeat(ToByte(&H0%), Length))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Module
