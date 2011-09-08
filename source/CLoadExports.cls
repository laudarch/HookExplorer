VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadExports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: David Zimmer 2002
'Site:      http://sandsprite.com
'
'Author:    David Zimmer <dzzie@yahoo.com>
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

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    name As Long
    base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public functions As New Collection  'of CExport
Public offset As CPEOffsets
Public errMessage As String


Function LoadExports(fpath As String) As Boolean
    
    Dim offsetExportTable As Long
    
    Set offset = New CPEOffsets
    Set functions = New Collection
    
    If Not offset.LoadFile(fpath) Then
        errMessage = offset.errMessage
        Exit Function
    End If
    
    If offset.RvaExportDirectory < 1 Then
        errMessage = "RvaExportDirectory = 0"
        Exit Function
    End If

    Dim ied As IMAGE_EXPORT_DIRECTORY
    Dim firstOrdial As Long
    Dim oset As Long
    Dim pointer As Long
    Dim i As Long
    Dim fhandle As Long
    
    Dim func() As Long
    Dim names() As String
    Dim ordials() As Integer '2bytes each
    
    'ntdll sent me for a loop :(
    'Debug.Print "RvaExportDir: " & Hex(offset.RvaExportDirectory)
    
    offsetExportTable = offset.RvaToOffset(offset.RvaExportDirectory)
    
    'Debug.Print "File Offset ExportDir: " & Hex(offsetExportTable)
    
    If offsetExportTable = 0 Then
        errMessage = "offsetExportTable = 0"
        Exit Function
    End If
    
    offsetExportTable = offsetExportTable + 1 '0 based file offsets
    
    fhandle = FreeFile
    Open fpath For Binary Access Read As fhandle
    
    Get fhandle, offsetExportTable, ied
    firstOrdial = ied.base
    oset = offset.RvaToOffset(ied.AddressOfFunctions) + 1
        
    'Debug.Print "AdrofFunc: " & Hex(ied.AddressOfFunctions)
    'Debug.Print "FileOffset:" & Hex(oset)
        
    For i = 0 To ied.NumberOfFunctions - 1
        Get fhandle, (oset + (i * 4)), pointer
        'If i = 0 Then
        '    Debug.Print "FuncPointer 0: " & Hex(pointer)
        'End If
        push func(), pointer
        'Debug.Print "Fx Addr " & i & ":" & Hex(pointer)
    Next
    
    'Debug.Print "AddofNames: " & Hex(ied.AddressOfNames)
    
    oset = offset.RvaToOffset(ied.AddressOfNames) + 1
    
    'Debug.Print "FileOffset: " & Hex(oset)
    
    Dim ns As String
    For i = 0 To ied.NumberOfNames - 1
        Get fhandle, (oset + (i * 4)), pointer
        'If i = 0 Then
        '    Debug.Print "Name pointer rva: " & Hex(pointer)
        '    Debug.Print "Name pointer foffset: " & Hex(offset.RvaToOffset(pointer))
        'End If
        ns = ExtractExportName(fhandle, offset.RvaToOffset(pointer))
        'If i = 0 Then
        '    Debug.Print "Name: " & ns
        'End If
        push names(), ns
    Next

    If ied.NumberOfNames > 0 Then
        Dim hasOrdials As Boolean
        Dim nameUb As Long, n As String
        
        nameUb = ied.NumberOfNames - 1
        hasOrdials = IIf(ied.AddressOfNameOrdinals > 0, True, False)
        
        If hasOrdials Then
            ReDim ordials(nameUb)
            Get fhandle, offset.RvaToOffset(ied.AddressOfNameOrdinals) + 1, ordials
        End If
        
        For i = 0 To nameUb
           
           'old method was simply:
           'If i > nameUb Then n = Empty Else n = names(i)
           'AddExport func(i), n, (ordials(i) + firstOrdial)
           'except failed for ntdll whos first ordial is #6 and 1st name offset is 305!
           
           n = names(i)
           If hasOrdials And ordials(i) <= UBound(func) Then 'sanity check for new calc
                AddExport func(ordials(i)), n, (ordials(i) + firstOrdial)
           Else
                'Stop
                'Debug.Print "Export Weirdness " & fpath & " i=" & i & " ordial(i)=" & ordials(i)
                AddExport func(i), "--> " & n, (ordials(i) + firstOrdial)
           End If
           
        Next
    End If
    
    OrdialSortCollection
    
    LoadExports = True
    
End Function

Private Function OrdialSortCollection()
    
    Dim x As CExport
    Dim tmp As New Collection
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To &HFFF
        If objKeyExistsInCollection(functions, "ordial:" & i) Then
            Set x = functions("ordial:" & i)
            If Not x Is Nothing Then
                tmp.Add x, "ordial:" & i
                functions.Remove "ordial:" & i
            End If
        End If
    Next
    
    For Each x In functions
        If Len(x.FunctionName) > 0 Then tmp.Add x
    Next
    
    Set functions = tmp
    
    
Exit Function
hell: DoEvents
    
End Function

Private Function ExtractExportName(f As Long, pointer As Long) As String
     Dim b As Byte
     Dim tmp As String
     
        If pointer <= 0 Or pointer >= LOF(f) Then
            ExtractExportName = "[No Name]"
        Else
            Get f, (pointer + 1), b
            While b <> 0
                tmp = tmp & Chr(b)
                Get f, , b
            Wend
            ExtractExportName = tmp
        End If
        
End Function


Private Sub AddExport(address As Long, name As String, Optional ordial As Long)
    Dim m  As New CExport
    m.FunctionAddress = address
    m.FunctionName = name
    m.FunctionOrdial = ordial
    On Error GoTo hell
    If Not objKeyExistsInCollection(functions, "ordial:" & ordial) Then
        functions.Add m, "ordial:" & ordial
    Else
        If ordial = 0 Then functions.Add m
    End If
    
    Exit Sub
hell:     DoEvents
End Sub



Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Function objKeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    objKeyExistsInCollection = True
 Exit Function
nope: objKeyExistsInCollection = False
End Function
