'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Linq

'This module contains the Microsoft EXEPACK related procedures.
Public Module EXEPackModule
   'This enumaration list the locations of the EXEPACK variables.
   Private Enum EXEPACKE
      DestinationLength = &HC%         'The unpacker code's destination in paragraphs relative to the start of the executable in memory.
      MemorySegment = &H4%             'The first memory segment containing the executable.
      RealCodeSegment = &H2%           'The executable's real initial code segment.
      RealInstructionPointer = &H0%    'The executable's real initial instruction pointer.
      RealStackPointer = &H8%          'The executable's real initial stack pointer.
      RealStackSegment = &HA%          'The executable's real initial stack segment.
      Signature = &H10%                'The unpacker's signature.
      SkipLength = &HE%                'The zero based number of paragraphs between the packed executable and the unpacker's variables.
      UnpackerSize = &H6%              'The unpacker's (variables + code + error message + packed relocation table) size specified in bytes.
   End Enum

   'This structure defines the EXEPACK variables.
   Private Structure EXEPACKStr
      Public RealInstructionPointer As Integer    'Defines the executable's real initial instruction pointer.
      Public RealCodeSegment As Integer           'Defines the executable's real initial code segment.
      Public MemorySegment As Integer             'Defines the first memory segment containing the executable.
      Public UnpackerSize As Integer              'Defines the unpacker's (variables + code + error message + packed relocation table) size specified in bytes.
      Public RealStackPointer As Integer          'Defines the executable's real initial stack pointer.
      Public RealStackSegment As Integer          'Defines the executable's real initial stack segment.
      Public DestinationLength As Integer         'Defines the unpacker code's destination in paragraphs relative to the start of the executable in memory.
      Public SkipLenth As Integer                 'Defines the zero based number of paragraphs between the packed executable and the unpacker's variables.
      Public Signature As Integer                 'Defines the unpacker's signature.
   End Structure

   Private Const PACKED_FILE_ERROR As String = "Packed file is corrupt"    'Contains the unpacker's error message.
   Private ReadOnly EXEPACK_PADDING As Char = ToChar(&HFF%)                'Contains the unpacker's padding byte.

   'This procedure creates an unpacked executable using the specified data and returns the result.
   Private Function CreateExecutable(MSDOSHeader As MSDOSHeaderStr, EXEPackHeader As EXEPACKStr, UnpackedData As List(Of Byte), RelocationTable As List(Of Byte)) As List(Of Byte)
      Try
         Dim EndOfRelocationTable As Integer = MSDOS_HEADER_SIZE + RelocationTable.Count
         Dim Executable As New List(Of Byte)
         Dim HeaderSize As Integer = CInt((((EndOfRelocationTable \ PAGE_SIZE) + &H1%) * PAGE_SIZE) / PARAGRAPH_SIZE)
         Dim PaddingSize As Integer = (HeaderSize * PARAGRAPH_SIZE) - EndOfRelocationTable
         Dim ExecutableSize As Integer = EndOfRelocationTable + PaddingSize + UnpackedData.Count

         With Executable
            .AddRange(ToLittleEndian(MSDOS_EXECUTABLE_SIGNATURE, &H2%))
            .AddRange(ToLittleEndian(ExecutableSize Mod PAGE_SIZE, &H2%))
            .AddRange(ToLittleEndian(CInt(ExecutableSize / PAGE_SIZE), &H2%))
            .AddRange(ToLittleEndian(CInt(RelocationTable.Count / &H4%), &H2%))
            .AddRange(ToLittleEndian(HeaderSize, &H2%))
            .AddRange(ToLittleEndian(MSDOSHeader.MinimumParagraphs, &H2%))
            .AddRange(ToLittleEndian(MSDOSHeader.MaximumParagraphs, &H2%))
            .AddRange(ToLittleEndian(EXEPackHeader.RealStackSegment, &H2%))
            .AddRange(ToLittleEndian(EXEPackHeader.RealStackPointer, &H2%))
            .AddRange(ToLittleEndian(MSDOSHeader.Checksum, &H2%))
            .AddRange(ToLittleEndian(EXEPackHeader.RealInstructionPointer, &H2%))
            .AddRange(ToLittleEndian(EXEPackHeader.RealCodeSegment, &H2%))
            .AddRange(ToLittleEndian(MSDOS_HEADER_SIZE, &H2%))
            .AddRange(ToLittleEndian(MSDOSHeader.OverlayNumber, &H2%))
            .AddRange(RelocationTable)
            .AddRange(TEXT_TO_BYTES(New String(PADDING, PaddingSize)))
            .AddRange(UnpackedData)
         End With

         Return Executable
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure unpacks the specfied packed data.
   Private Function Unpack(PackedData As List(Of Byte)) As List(Of Byte)
      Try
         Dim Count As New Integer
         Dim FillByte As New Byte
         Dim Opcode As New Integer
         Dim Position As Integer = 0
         Dim UnpackedData As New List(Of Byte)

         Do While PackedData(Position) = ToByte(EXEPACK_PADDING)
            Position += &H1%
         Loop

         Do
            Opcode = PackedData(Position)
            Position += &H1%
            Count = (CInt(PackedData(Position)) << &H8%) Or PackedData(Position + &H1%)
            Position += &H2%

            If (Opcode And &HFE%) = &HB0% Then
               FillByte = ToByte(PackedData(Position))

               For Index As Integer = &H0% To Count - &H1%
                  UnpackedData.Add(FillByte)
               Next Index

               Position += &H1%
            ElseIf (Opcode And &HFE%) = &HB2% Then
               For Index As Integer = &H0% To Count - &H1%
                  UnpackedData.Add(PackedData(Position))
                  Position += &H1%
               Next Index
            Else
               Return Nothing
            End If
         Loop Until (Opcode And &H1%) = &H1%

         Do While Position < PackedData.Count
            UnpackedData.Add(PackedData(Position))
            Position += &H1%
         Loop

         Return UnpackedData
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure gives the command to unpack the specified packed executable and returns the result.
   Public Function UnpackExecutable(FileData As List(Of Byte)) As List(Of Byte)
      Try
         Dim EXEPack As New EXEPACKStr
         Dim EXEPackOffset As New Integer
         Dim MSDOSHeader As New MSDOSHeaderStr
         Dim PackedData As New List(Of Byte)
         Dim PackedDataStart As New Integer
         Dim ReversedData As New List(Of Byte)
         Dim UnpackedData As New List(Of Byte)

         With MSDOSHeader
            .Signature = GET_WORD(FileData, MSDOSHeaderE.Signature)
            .ImageSizeModulo = GET_WORD(FileData, MSDOSHeaderE.ImageSizeModulo)
            .ImageSize = GET_WORD(FileData, MSDOSHeaderE.ImageSize)
            .RelocationCount = GET_WORD(FileData, MSDOSHeaderE.RelocationCount)
            .HeaderSize = GET_WORD(FileData, MSDOSHeaderE.HeaderSize)
            .MinimumParagraphs = GET_WORD(FileData, MSDOSHeaderE.MinimumParagraphs)
            .MaximumParagraphs = GET_WORD(FileData, MSDOSHeaderE.MaximumParagraphs)
            .StackSegment = GET_WORD(FileData, MSDOSHeaderE.StackSegment)
            .StackPointer = GET_WORD(FileData, MSDOSHeaderE.StackPointer)
            .Checksum = GET_WORD(FileData, MSDOSHeaderE.Checksum)
            .InstructionPointer = GET_WORD(FileData, MSDOSHeaderE.InstructionPointer)
            .CodeSegment = GET_WORD(FileData, MSDOSHeaderE.CodeSegment)
            .RelocationTableOffset = GET_WORD(FileData, MSDOSHeaderE.RelocationTableOffset)
            .OverlayNumber = GET_WORD(FileData, MSDOSHeaderE.OverlayNumber)

            EXEPackOffset = (.HeaderSize + .CodeSegment) * PARAGRAPH_SIZE
            PackedDataStart = .HeaderSize * PARAGRAPH_SIZE
         End With

         With EXEPack
            .RealInstructionPointer = GET_WORD(FileData, EXEPackOffset + EXEPACKE.RealInstructionPointer)
            .RealCodeSegment = GET_WORD(FileData, EXEPackOffset + EXEPACKE.RealCodeSegment)
            .MemorySegment = GET_WORD(FileData, EXEPackOffset + EXEPACKE.MemorySegment)
            .UnpackerSize = GET_WORD(FileData, EXEPackOffset + EXEPACKE.UnpackerSize)
            .RealStackPointer = GET_WORD(FileData, EXEPackOffset + EXEPACKE.RealStackPointer)
            .RealStackSegment = GET_WORD(FileData, EXEPackOffset + EXEPACKE.RealStackSegment)
            .DestinationLength = GET_WORD(FileData, EXEPackOffset + EXEPACKE.DestinationLength)
            .SkipLenth = GET_WORD(FileData, EXEPackOffset + EXEPACKE.SkipLength)
            .Signature = GET_WORD(FileData, EXEPackOffset + EXEPACKE.Signature)
         End With

         ReversedData.AddRange(GetBytes(FileData, &H0%, PackedDataStart))
         PackedData = GetBytes(FileData, PackedDataStart, EXEPackOffset - PackedDataStart)
         PackedData.Reverse()
         ReversedData.AddRange(PackedData)
         UnpackedData = Unpack(GetBytes(ReversedData, PackedDataStart, ReversedData.Count - PackedDataStart))
         UnpackedData.Reverse()

         Return CreateExecutable(MSDOSHeader, EXEPack, UnpackedData, UnpackRelocationTable(FileData, MSDOSHeader, EXEPack))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure unpacks the specified packed relocation table and returns the result.
   Private Function UnpackRelocationTable(FileData As List(Of Byte), MSDOSHeader As MSDOSHeaderStr, EXEPackHeader As EXEPACKStr) As List(Of Byte)
      Try
         Dim Count As Integer = &H0%
         Dim Entry As New Integer
         Dim EXEPackOffset As Integer = (MSDOSHeader.HeaderSize + MSDOSHeader.CodeSegment) * PARAGRAPH_SIZE
         Dim PackedRelocationTable As New List(Of Byte)
         Dim Position As New Integer
         Dim RelocationOffset As Integer = FindBytes(FileData, TEXT_TO_BYTES(PACKED_FILE_ERROR), EXEPackOffset)
         Dim RelocationTableSize As Integer = CInt((((EXEPackHeader.UnpackerSize + EXEPackOffset + PACKED_FILE_ERROR.Length) - (PARAGRAPH_SIZE * &H4%)) / &H2%) * &H4%)
         Dim UnpackedRelocationTable As New List(Of Byte)

         PackedRelocationTable = New List(Of Byte)(GetBytes(FileData, RelocationOffset, FileData.Count - RelocationOffset))

         Position = PACKED_FILE_ERROR.Length
         For PackedRelocationItem As Integer = &H0% To &HF%
            Count = GET_WORD(PackedRelocationTable, Position)
            Position += &H2%
            For UnpackedRelocationItem As Integer = &H0% To Count - &H1%
               Entry = GET_WORD(PackedRelocationTable, Position)
               Position += &H2%
               UnpackedRelocationTable.AddRange(ToLittleEndian(Entry, &H2%))
               UnpackedRelocationTable.AddRange(ToLittleEndian(PackedRelocationItem << &HC%, &H2%))
            Next UnpackedRelocationItem
         Next PackedRelocationItem

         Return UnpackedRelocationTable
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module

