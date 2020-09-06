'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Environment
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

'This class contains the movie related procedures.
Public Class CartoonersClass
   Inherits DataFileClass

   'This enumeration defines the data chunk information.
   Private Enum DataChunkE As Integer
      Description   'A data chunk's description.
      IsBinary      'Indicates whether a chunk is raw binary data.
      Offset        'A data chunk's offset.
      Length        'A data chunk's length.
   End Enum

   'This structure defines the description, offset and size of data chunks used by Cartooners.
   Private Structure DataChunkStr
      Public Description As String   'Defines a data chunk's description.
      Public IsBinary As Boolean     'Indicates whether a chunk is raw binary data.
      Public Offset As Integer       'Defines a data chunk's offset.
      Public Length As Integer       'Defines a data chunk's length.
   End Structure

   'This enumeration contains the locations of all icons used by Cartooners.
   Private Enum IconLocationsE As Integer
      CheckedCircle = &H312EC    'The checked circle icon.
      Circle = &H312A4%          'The circle icon.
      CrossArrows = &H31750%     'The cross arrows icon.
      Eraser = &H315A0%          'The eraser icon.
      FilledCircle = &H31334%    'The filled circle icon.
      LeftArrow = &H31438%       'The left arrow icon.
      LoopArrow = &H316C0%       'The loop arrow icon.
      RightArrow = &H314EC       'The right arrow icon.
      SpeechBalloon = &H31630%   'The speech balloon icon.
   End Enum

   'This enumeration contains the location of all images used by Cartooners.
   Private Enum ImageLocationsE As Integer
      MenuScreen = &H29A60%      'The run-length compressed menu screen image.
      TicketScreen = &H2F560%    'The run-length compressed ticket screen image.
      TitleScreen = &H2CB82%     'The run-length compressed title screen image.
   End Enum

   'This enumeration contains the location of all palettes used by Cartooners.
   Private Enum PaletteLocationsE As Integer
      GUICGAEGA = &H31E50%          'The GUI's CGA/EGA palette.
      MenuScreenCGA = &H29A20%      'The menu screen's CGA palette.
      MenuScreenEGA = &H29A40%      'The menu screen's EGA palette.
      TicketScreenCGA = &H2F540%    'The ticket screen CGA palette.
      TicketScreenEGA = &H2F520%    'The ticket screen EGA palette.
      TitleScreenCGA = &H2CB40%     'The title screen's CGA palette.
      TitleScreenEGA = &H2CB60%     'The title screen's EGA palette.
   End Enum

   'This enumeration contains a list of the palettes used by Cartooners.
   Private Enum PalettesE As Integer
      GUICGAEGA = &H6%         'The GUI's CGA/EGA palette.
      MenuScreenCGA = &H0%     'The menu screen's CGA palette.
      MenuScreenEGA = &H1%     'The menu screen's EGA palette.
      TicketScreenCGA = &H5%   'The ticket screen's CGA palette.
      TicketScreenEGA = &H4%   'The ticket screen's EGA palette.
      TitleScreenCGA = &H2%    'The title screen's CGA palette.
      TitleScreenEGA = &H3%    'The title screen's EGA palette.
   End Enum

   Private Const BYTES_PER_ROW As Integer = &HA0%                  'Contains the number of bytes per pixel row.
   Private Const DATA_CHUNK_DELIMITER As Char = ";"c               'Contains the data chunk information delimiter.
   Private Const EXPECTED_NAME As String = "Cartoons.exe"          'Contains the Cartooners executable's expected file name.
   Private Const EXPECTED_PACKED_SIZE As Integer = &H36A5F%        'Contains the Cartooners executable's expected packed file size.
   Private Const EXPECTED_UNPACKED_SIZE As Integer = &H39F20%      'Contains the Cartooners executable's expected unpacked file size.
   Private Const SCREEN_HEIGHT As Integer = &HC8%                  'Contains the screen height used by Cartooners in pixels.
   Private Const SCREEN_WIDTH As Integer = &H140%                  'Contains the screen width used by Cartooners in pixels.

   Private ReadOnly IMAGE_SIZES As New Dictionary(Of Integer, Integer) From
      {{ImageLocationsE.MenuScreen, &H262C%},
      {ImageLocationsE.TicketScreen, &H1A9C%},
      {ImageLocationsE.TitleScreen, &H2995%}}   'The image sizes.

   'The menu items used by this class.
   Private WithEvents DisplayDataMenu As New ToolStripMenuItem With {.Text = "Display &Data"}
   Private WithEvents DisplayDataSubmenu As New ToolStripComboBox
   Private WithEvents DisplayInformationMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F1, .Text = "Display &Information"}
   Private WithEvents DisplayPalettesMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F2, .Text = "Display &Palettes"}

   'This procedure initializes this class.
   Public Sub New(PathO As String, Optional DataFileMenu As ToolStripMenuItem = Nothing)
      Try
         Dim TextSubMenuItems As New List(Of String)

         If DataFile(CartoonersPath:=PathO).Data IsNot Nothing AndAlso DataFileMenu IsNot Nothing Then
            For Each DataChunk As DataChunkStr In DataChunks()
               TextSubMenuItems.Add(DataChunk.Description)
            Next DataChunk
            TextSubMenuItems.Sort()

            With DisplayDataSubmenu
               .Items.Clear()
               TextSubMenuItems.ForEach(Sub(SubMenuItem As String) .Items.Add(New ToolStripMenuItem With {.CheckOnClick = True, .Text = SubMenuItem}))
            End With

            DisplayDataMenu.DropDownItems.Add(DisplayDataSubmenu)

            With DataFileMenu
               .DropDownItems.Clear()
               .DropDownItems.AddRange({DisplayInformationMenu, DisplayPalettesMenu, DisplayDataMenu})
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
         Dim Items() As String = {}
         Dim Lines As List(Of String) = Nothing
         Static CurrentDataChunks As List(Of DataChunkStr) = Nothing

         If CurrentDataChunks Is Nothing Then
            CurrentDataChunks = New List(Of DataChunkStr)

            Lines = New List(Of String)(My.Resources.Data_Cartooners_Chunks.Split({NewLine}, StringSplitOptions.None))
            Lines.RemoveAt(0)
            For Each Line As String In Lines
               If Not Line.Trim() = Nothing Then
                  Items = Line.Split(DATA_CHUNK_DELIMITER)
                  CurrentDataChunks.Add(New DataChunkStr With {.Description = Items(DataChunkE.Description), .IsBinary = CBool(Items(DataChunkE.IsBinary)), .Offset = CInt(Items(DataChunkE.Offset)), .Length = CInt(Items(DataChunkE.Length))})
               End If
            Next Line
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
                        If BitConverter.ToInt16(.Data.ToArray(), MSDOSHeaderE.Signature) = MSDOS_EXECUTABLE_SIGNATURE Then
                           .Path = CartoonersPath
                           EXEHeaderSize(NewData:= .Data)
                           Palettes(Refresh:=True)
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
         Dim IsBinary As New Boolean
         Dim Length As New Integer
         Dim NewText As New StringBuilder
         Dim Offset As New Integer

         For Index As Integer = 0 To DataChunks().Count - 1
            If DataChunks(Index).Description = DisplayDataSubmenu.Text Then
               IsBinary = DataChunks(Index).IsBinary
               Length = DataChunks(Index).Length
               Offset = DataChunks(Index).Offset + EXEHeaderSize()
            End If
         Next Index

         NewText.Append($"[{Description}]{NewLine}")
         If IsBinary Then
            NewText.Append($"{Escape(GetBytes(DataFile().Data, Offset, Length), " "c, EscapeAll:=True).Trim()}")
         Else
            NewText.Append(Escape(ConvertMSDOSLineBreak(GetString(DataFile.Data, Offset, Length).Replace(DELIMITER, NewLine))))
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

   'This procedure displays the current Cartooners executable's information palettes.
   Private Sub DisplayPalettesMenu_Click(sender As Object, e As EventArgs) Handles DisplayPalettesMenu.Click
      Try
         Dim Descriptions As New List(Of String)

         For Each Location As Integer In CType([Enum].GetValues(GetType(PaletteLocationsE)), Integer())
            Descriptions.Add(DirectCast(Location + EXEHeaderSize(), PaletteLocationsE).ToString())
         Next Location

         UpdateDataBox(GBRToText("Cartooners' palettes:", Palettes(), Descriptions))
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

         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.MenuScreen + EXEHeaderSize(), IMAGE_SIZES(ImageLocationsE.MenuScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.MenuScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Menu screen.png"), ImageFormat.Png)
         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.TicketScreen + EXEHeaderSize(), IMAGE_SIZES(ImageLocationsE.TicketScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.TicketScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Ticket screen.png"), ImageFormat.Png)
         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.TitleScreen + EXEHeaderSize(), IMAGE_SIZES(ImageLocationsE.TitleScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.TitleScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Title screen.png"), ImageFormat.Png)

         For Each Location As Integer In CType([Enum].GetValues(GetType(IconLocationsE)), Integer())
            IconHeight = BitConverter.ToInt16(DataFile().Data.ToArray(), Location + EXEHeaderSize() + &H2%)
            IconWidth = BitConverter.ToInt16(DataFile().Data.ToArray(), Location + EXEHeaderSize() + &H4%)
            BytesPerRow = If(IconWidth Mod PIXELS_PER_BYTE = 0, IconWidth \ PIXELS_PER_BYTE, (IconWidth + &H1%) \ PIXELS_PER_BYTE)
            IconSize = BytesPerRow * IconHeight
            Draw4BitImage(GetBytes(DataFile().Data, Location + EXEHeaderSize() + &H6%, IconSize), IconWidth, IconHeight, Palettes()(PalettesE.TitleScreenEGA), BytesPerRow).Save($"{Path.Combine(ExportPath, DirectCast(Location, IconLocationsE).ToString())}.png", ImageFormat.Png)
         Next Location

         Process.Start(New ProcessStartInfo With {.FileName = ExportPath, .WindowStyle = ProcessWindowStyle.Normal})
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the Cartooners executable's palettes.
   Private Function Palettes(Optional Refresh As Boolean = False) As List(Of List(Of Color))
      Try
         Static CurrentPalettes As New List(Of List(Of Color))

         If Refresh Then
            CurrentPalettes.Clear()

            For Each Location As Integer In CType([Enum].GetValues(GetType(PaletteLocationsE)), Integer())
               CurrentPalettes.Add(GBRPalette(DataFile().Data, Location + EXEHeaderSize()))
            Next Location
         End If

         Return CurrentPalettes
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Class
