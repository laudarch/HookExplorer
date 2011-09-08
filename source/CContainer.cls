VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Public Entries As New Collection 'of CEntry 'all of them
Public AllHookedEntries As New Collection 'of CEntry  'only hooked (includes self hookers)
Public RealHookedEntries As New Collection 'only external module hooks
Public FilteredHookedEntries As New Collection

Public li As ListItem
Public DllName As String
Public ImportsScanned As Boolean
Public ExportsScanned As Boolean
Public ExportsTotal As Long
Public size As String
Public Module As CModule

Sub AddEntry(memadr As Long, name As String, adrValue As Long, _
             cmd As String, modName As String, Hooked As Boolean, _
             HookAdr As Long, HookMod As String)
    
    Dim e As New CEntry
    Dim dll
    
    With e
        .MemAddress = memadr
        .name = name
        .AddressValue = adrValue
        .FirstInstruction = cmd
        .ModuleName = modName
        .isHooked = Hooked
        .HookAddress = HookAdr
        .HookModule = HookMod
        If InStr(1, HookMod, modName, vbTextCompare) > 0 Then .SameDllHook = True
    End With
    
    If Hooked Then
    
        AllHookedEntries.Add e
        
        If Not e.SameDllHook Then
            RealHookedEntries.Add e
            
            For Each dll In Form1.IgnoreList
                If InStr(1, dll, HookMod, vbTextCompare) > 0 Then
                    GoTo skipIt
                End If
            Next
            FilteredHookedEntries.Add e
            
        End If
        
    End If

skipIt:
    Entries.Add e
End Sub

Sub ReapplyFilters()
    
    Dim e As New CEntry
    Dim dll
    
    Set FilteredHookedEntries = Nothing
    Set FilteredHookedEntries = New Collection
    
    For Each e In Entries
        If e.isHooked And Not e.SameDllHook Then
            For Each dll In Form1.IgnoreList
                If InStr(1, dll, e.HookModule, vbTextCompare) > 0 Then
                    GoTo skipIt
                End If
            Next
            FilteredHookedEntries.Add e
        End If
skipIt:
    Next
        
        
End Sub

Sub DumpSelectedToHandle(fHandle As Long)
    
    Dim c As Collection
    Dim e As CEntry
    Dim tmp As String
    
    On Error Resume Next
    
    Select Case Form1.DisplayOption
        Case doShowAll:    Set c = Entries
        Case doIgnoreList: Set c = FilteredHookedEntries
        Case doHideSelf:   Set c = RealHookedEntries
        Case doStandard:   Set c = AllHookedEntries
    End Select
     
    For Each e In c
            tmp = IIf(e.MemAddress < 0, "N/A", Hex(e.MemAddress))
            tmp = tmp & vbTab & Hex(e.AddressValue)
            tmp = tmp & vbTab & e.name
            tmp = tmp & vbTab & e.FirstInstruction
            
           'If e.isHooked And e.AddressValue > 0 Then
                If Form1.KeyExistsInCollection(Form1.Symbols, "adr:" & Hex(e.AddressValue)) Then
                    tmp = tmp & vbTab & Form1.Symbols("adr:" & Hex(e.AddressValue))
                    'li.SubItems(4) = Split(li.SubItems(4), ".")(1)
                Else
                    tmp = tmp & vbTab & "Unknown"
                End If
           'Else
           '     tmp = tmp & vbTab
           'End If
            
            tmp = tmp & vbTab & e.HookModule
            Print #fHandle, tmp
        Next
    
End Sub

Sub FillOutListView()
    
    Dim e As CEntry
    Dim li As ListItem
    Dim c As Collection
    
    On Error Resume Next
    
    Select Case Form1.DisplayOption
        Case doShowAll:    Set c = Entries
        Case doIgnoreList: Set c = FilteredHookedEntries
        Case doHideSelf:   Set c = RealHookedEntries
        Case doStandard:   Set c = AllHookedEntries
    End Select
    
    With Form1.lvImports.ListItems
         
        .Clear
        
        For Each e In c       'N/A means found from Export Scan no IAT mem adr
            Set li = .Add(, , IIf(e.MemAddress < 0, "N/A", Hex(e.MemAddress)))
            li.SubItems(1) = Hex(e.AddressValue)
            li.SubItems(2) = e.name
            li.SubItems(3) = e.FirstInstruction
            li.SubItems(5) = e.HookModule
            
            'If e.isHooked And e.AddressValue > 0 Then
                If Form1.KeyExistsInCollection(Form1.Symbols, "adr:" & Hex(e.AddressValue)) Then
                    li.SubItems(4) = Form1.Symbols("adr:" & Hex(e.AddressValue))
                    'li.SubItems(4) = Split(li.SubItems(4), ".")(1)
                Else
                    li.SubItems(4) = "Unknown"
                End If
            'End If
                    
        Next
    
    End With
    
    
    
End Sub




Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
