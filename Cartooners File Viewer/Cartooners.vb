'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

'This class contains the Cartooners related procedures.
Public Class CartoonersClass
   Inherits DataFileClass

   'This enumeration contains the data chunk properties.
   Private Enum DataChunkPropertiesE As Integer
      Description   'A data chunk's description.
      Type          'A data chunk's type.
      Offset        'A data chunk's offset.
      Related       'A related data chunk's offset.
      Length        'A data chunk's length.
   End Enum

   'This structure defines the description, offset and size of data chunks used by Cartooners.
   Private Structure DataChunkStr
      Public Description As String   'Defines a data chunk's description.
      Public Type As String          'Defines a data chunk's type.
      Public Offset As Integer       'Defines a data chunk's offset.
      Public Related As Integer      'Defines a related data chunk's offset.
      Public Length As Integer       'Defines a data chunk's length.
   End Structure

   Private Const BYTES_PER_ROW As Integer = &HA0%                  'Contains the number of bytes per pixel row.
   Private Const EXPECTED_NAME As String = "Cartoons.exe"          'Contains the Cartooners executable's expected file name.
   Private Const EXPECTED_PACKED_SIZE As Integer = &H36A5F%        'Contains the Cartooners executable's expected packed file size.
   Private Const EXPECTED_UNPACKED_SIZE As Integer = &H39F20%      'Contains the Cartooners executable's expected unpacked file size.
   Private Const SCREEN_HEIGHT As Integer = &HC8%                  'Contains the screen height used by Cartooners in pixels.
   Private Const SCREEN_WIDTH As Integer = &H140%                  'Contains the screen width used by Cartooners in pixels.

   Private ReadOnly DATA_CHUNK_PROPERTY_DELIMITER As Char = ToChar(9)   'Contains the data chunk property delimiter.

   'The menu items used by this class.
   Private WithEvents DisplayDataMenu As New ToolStripMenuItem With {.Text = "Display &Data"}
   Private WithEvents DisplayDataSubmenu As New ToolStripComboBox
   Private WithEvents DisplayInformationMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F1, .Text = "Display &Information"}

   'This procedure initializes this class.
   Public Sub New(PathO As String, Optional DataFileMenu As ToolStripMenuItem = Nothing)
      Try
         Dim DataChunkDescriptions As New List(Of String)

         If DataFile(CartoonersPath:=PathO).Data IsNot Nothing AndAlso DataFileMenu IsNot Nothing Then
            For Each DataChunk As DataChunkStr In DataChunks()
               DataChunkDescriptions.Add(DataChunk.Description)
            Next DataChunk
            DataChunkDescriptions.Sort()

            With DisplayDataSubmenu
               .Items.Clear()
               DataChunkDescriptions.ForEach(Sub(SubMenuItem As String) .Items.Add(New ToolStripMenuItem With {.CheckOnClick = True, .Text = SubMenuItem}))
            End With

            DisplayDataMenu.DropDownItems.Add(DisplayDataSubmenu)

            With DataFileMenu
               .DropDownItems.Clear()
               .DropDownItems.AddRange({DisplayInformationMenu, DisplayDataMenu})
               .Text = "&Cartooners"
               .Visible = True
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages Cartooner's data chunk information.
   Private Function DataChunks() As List(Of DataChunkStr)
      Try
         Dim Chunks As List(Of String) = Nothing
         Dim Properties() As String = {}
         Static CurrentDataChunks As List(Of DataChunkStr) = Nothing

         If CurrentDataChunks Is Nothing Then
            CurrentDataChunks = New List(Of DataChunkStr)

            Chunks = New List(Of String)(My.Resources.Cartooners_Executable.Split({NewLine}, StringSplitOptions.None))
            Chunks.RemoveAt(0)
            For Each Chunk As String In Chunks
               If Not Chunk.Trim() = Nothing Then
                  Properties = Chunk.Split(DATA_CHUNK_PROPERTY_DELIMITER)
                  CurrentDataChunks.Add(New DataChunkStr With {.Description = Properties(DataChunkPropertiesE.Description), .Type = Properties(DataChunkPropertiesE.Type).Trim().ToLower(), .Offset = CInt(Properties(DataChunkPropertiesE.Offset)), .Related = CInt(Properties(DataChunkPropertiesE.Related)), .Length = CInt(Properties(DataChunkPropertiesE.Length))})
               End If
            Next Chunk
         End If

         Return CurrentDataChunks
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedures manages the Cartooners executable's data file.
   Private Function DataFile(Optional CartoonersPath As String = Nothing) As DataFileStr
      Try
         Static CurrentFile As New DataFileStr With {.Data = Nothing, .Path = Nothing}

         With CurrentFile
            If Not CartoonersPath = Nothing Then
               If Path.GetFileName(CartoonersPath).ToUpper() = EXPECTED_NAME.ToUpper() Then
                  If New FileInfo(CartoonersPath).Length = EXPECTED_PACKED_SIZE Then
                     .Data = UnpackExecutable(New List(Of Byte)(File.ReadAllBytes(CartoonersPath)))
                     If .Data.Count = EXPECTED_UNPACKED_SIZE Then
                        If BitConverter.ToUInt16(.Data.ToArray(), MSDOSHeaderE.Signature) = MSDOS_EXECUTABLE_SIGNATURE Then
                           .Path = CartoonersPath
                           EXEHeaderSize(NewData:= .Data)
                           RelocationItems(NewData:= .Data)
                           DisplayInformationMenu.PerformClick()
                        Else
                           MessageBox.Show("Invalid MS-DOS executable signature.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                           .Data = Nothing
                        End If
                     Else
                        MessageBox.Show($"An error occurred during unpacking. Expected unpacked size: {EXPECTED_UNPACKED_SIZE} bytes. Actual unpacked size: { .Data.Count} bytes.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        .Data = Nothing
                     End If
                  Else
                     MessageBox.Show($"Wrong size. Expected packed size: {EXPECTED_PACKED_SIZE} bytes.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                  End If
               Else
                  MessageBox.Show($"Wrong Cartooners executable name. Expected Cartooners executable name: ""{EXPECTED_NAME}"".", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
               End If
            End If
         End With

         Return CurrentFile
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the current Cartooners executable's speech balloon slots/way menu.
   Private Sub DisplayDataSubmenu_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DisplayDataSubmenu.SelectedIndexChanged
      Try
         Dim Description As String = DisplayDataSubmenu.Text
         Dim IconHeight As New Integer
         Dim IconType As New Integer
         Dim IconWidth As New Integer
         Dim Length As New Integer
         Dim NewText As New StringBuilder
         Dim Palettes As List(Of List(Of Color)) = Nothing
         Dim Offset As New Integer
         Dim Type As String = Nothing

         For Index As Integer = 0 To DataChunks().Count - 1
            If DataChunks(Index).Description = DisplayDataSubmenu.Text Then
               Type = DataChunks(Index).Type
               Length = DataChunks(Index).Length
               Offset = DataChunks(Index).Offset + EXEHeaderSize()
            End If
         Next Index

         If Offset + Length > DataFile().Data.Count Then
            MessageBox.Show($"Attempting to read {(Offset + Length) - DataFile().Data.Count} byte(s) beyond of the end of the available data at position {Offset}.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Length = DataFile().Data.Count - Offset
         End If

         NewText.Append($"[{Description}] ({Type}){NewLine}")
         If Type = "binary" OrElse Type = "image" Then
            NewText.Append($"{Escape(GetBytes(DataFile().Data, Offset, Length), " "c, EscapeAll:=True).Trim()}")
         ElseIf Type = "icon" Then
            IconHeight = BitConverter.ToUInt16(DataFile().Data.ToArray(), Offset + &H2%)
            IconType = BitConverter.ToUInt16(DataFile().Data.ToArray(), Offset)
            IconWidth = BitConverter.ToUInt16(DataFile().Data.ToArray(), Offset + &H4%)

            NewText.Append($"Size: {IconWidth * 2} x {IconHeight} - Type: {IconType} {NewLine}{NewLine}")
            NewText.Append($"{Escape(GetBytes(DataFile().Data, Offset, Length), " "c, EscapeAll:=True).Trim()}")
         ElseIf Type = "palette" Then
            Palettes = New List(Of List(Of Color))
            Palettes.Add(New List(Of Color)(GBRPalette(DataFile().Data, Offset)))
            NewText.Append(GBRToText(, Palettes))
         ElseIf Type = "text" Then
            NewText.Append(Escape(ConvertMSDOSLineBreak(GetString(DataFile.Data, Offset, Length)).Replace(DELIMITER, NewLine)))
         End If

         UpdateDataBox(NewText.ToString())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current Cartooners executable's information.
   Private Sub DisplayInformationMenu_Click(sender As Object, e As EventArgs) Handles DisplayInformationMenu.Click
      Try
         Dim NewText As New StringBuilder

         With DataFile()
            NewText.Append($"{ .Path}:{NewLine}")
            NewText.Append($"File size: {New FileInfo(.Path).Length} bytes.{NewLine}{NewLine}")
            NewText.Append(GetEXEHeaderInformation(.Data))

            UpdateDataBox(NewText.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure exports the Cartooners executable's images.
   Public Overloads Sub Export(ExportPath As String)
      Try
         Dim BytesPerRow As New Integer
         Dim IconHeight As New Integer
         Dim IconSize As New Integer
         Dim IconWidth As New Integer

         ExportPath = Path.Combine(ExportPath, Path.GetFileNameWithoutExtension("Cartooners Executable"))
         Directory.CreateDirectory(ExportPath)

         File.WriteAllBytes(Path.Combine(ExportPath, "Cartoons.Unpacked.exe"), DataFile().Data.ToArray())

         For Each Chunk As DataChunkStr In DataChunks()
            With Chunk
               If .Type = "icon" Then
                  IconHeight = BitConverter.ToUInt16(DataFile().Data.ToArray(), .Offset + EXEHeaderSize() + &H2%)
                  IconWidth = BitConverter.ToUInt16(DataFile().Data.ToArray(), .Offset + EXEHeaderSize() + &H4%)
                  BytesPerRow = If(IconWidth Mod PIXELS_PER_BYTE = 0, IconWidth \ PIXELS_PER_BYTE, (IconWidth + &H1%) \ PIXELS_PER_BYTE)
                  IconSize = BytesPerRow * IconHeight
                  Draw4BitImage(GetBytes(DataFile().Data, .Offset + EXEHeaderSize() + &H6%, IconSize), IconWidth, IconHeight, GBRPalette(DataFile().Data, .Related + EXEHeaderSize()), BytesPerRow).Save($"{Path.Combine(ExportPath, .Description)}.png", ImageFormat.Png)
               ElseIf .Type = "image" Then
                  Draw4BitImage(DecompressRLE(DataFile().Data, .Offset + EXEHeaderSize(), .Length), SCREEN_WIDTH, SCREEN_HEIGHT, GBRPalette(DataFile().Data, .Related + EXEHeaderSize()), BYTES_PER_ROW).Save($"{Path.Combine(ExportPath, .Description)}.png", ImageFormat.Png)
               End If
            End With
         Next Chunk

         ExportMap(ExportPath)

         Process.Start(New ProcessStartInfo With {.FileName = ExportPath, .WindowStyle = ProcessWindowStyle.Normal})
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure exports a map of the Cartooner's executable.
   Private Sub ExportMap(ExportPath As String)
      Try
         Dim CurrentChunk As DataChunkStr? = Nothing
         Dim Data As New List(Of Byte)(DataFile().Data)
         Dim Map As New StringBuilder()
         Dim Position As Integer = &H0%
         Dim RelocationItemPositions As New List(Of Integer)

         RelocationItems().ForEach(Sub(Item As SegmentOffsetStr) RelocationItemPositions.Add(Item.Segment << &H4% Or Item.Offset))

         Data.RemoveRange(&H0%, EXEHeaderSize())
         Do Until Position >= Data.Count - &H1%
            CurrentChunk = Nothing
            For Each Chunk As DataChunkStr In DataChunks()
               If Position = Chunk.Offset Then
                  CurrentChunk = Chunk
                  Exit For
               End If
            Next Chunk

            If CurrentChunk Is Nothing Then
               If RelocationItemPositions.Contains(Position) Then Map.Append("*"c)
               Map.Append($"{Data(Position):X2} ")
               Position += &H1%
            Else
               With CurrentChunk.Value
                  Map.Append($"{NewLine}{NewLine}[BEGIN { .Description}]{NewLine}")
                  For SubPosition As Integer = .Offset To (.Offset + .Length) - &H1%
                     If RelocationItemPositions.Contains(SubPosition) Then Map.Append("*"c)
                     Map.Append($"{Data(SubPosition):X2} ")
                     Position += &H1%
                  Next SubPosition
                  Map.Append($"{NewLine}[END { .Description}]{NewLine}{NewLine}")
               End With
            End If
         Loop

         File.WriteAllText(Path.Combine(ExportPath, "Cartooners Executable Map.txt"), Map.ToString())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Class
