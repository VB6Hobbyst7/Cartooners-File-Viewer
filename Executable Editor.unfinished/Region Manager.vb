'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.IO
Imports System.Text

'This module contains the region manager procedures.
Public Module RegionManagerModule
   'This class defines the regions for an executable's data.
   Public Class RegionClass
      Public Description As String   'Defines a region's description.
      Public Offset As Integer       'Defines a region's offset.
      Public Length As Integer       'Defines a region's length.
      Public EndO As Integer         'Defines a region's end.
   End Class

   'This enumeration lists the regions' properties.
   Public Enum RegionPropertiesE As Integer
      Description   'A region's description.
      Offset        'A region's offset.
      Length        'A region's length.
      EndO          'A region's end.
   End Enum

   Public ReadOnly REGION_PROPERTY_DELIMITER As Char = ControlChars.Tab   'Defines the region property delimiter.

   'This procedure checks whether the current regions overlap and returns the result.
   Public Function CheckForRegionOverlap() As String
      Try
         Dim Overlaps As New StringBuilder

         For Each Region As RegionClass In Regions()
            For Each OtherRegion As RegionClass In Regions()
               If Region IsNot OtherRegion AndAlso Region.EndO > OtherRegion.Offset AndAlso Region.EndO <= OtherRegion.EndO Then
                  Overlaps.Append($"""{Region.Description} "" overlaps with ""{OtherRegion.Description}"".{NewLine}")
               End If
            Next OtherRegion
         Next Region

         Return Overlaps.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the current executable regions.
   Public Function Regions(Optional RegionPath As String = Nothing) As List(Of RegionClass)
      Try
         Dim Properties() As String = {}
         Static CurrentRegions As List(Of RegionClass) = Nothing

         If Not RegionPath = Nothing Then
            CurrentRegions = New List(Of RegionClass)

            For Each RegionLine As String In File.ReadAllLines(RegionPath)
               If Not RegionLine.Trim() = Nothing Then
                  Properties = RegionLine.Split(REGION_PROPERTY_DELIMITER)
                  CurrentRegions.Add(New RegionClass With {.Description = Properties(RegionPropertiesE.Description), .Offset = CInt(Properties(RegionPropertiesE.Offset)), .Length = CInt(Properties(RegionPropertiesE.Length)), .EndO = .Offset + .Length})
               End If
            Next RegionLine
         End If

         Return CurrentRegions
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
