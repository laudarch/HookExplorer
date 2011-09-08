VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadImports"
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

Private fhandle As Long 'my FileHandle

Private Type IMAGE_IMPORT_DIRECTORY
    pFuncAry As Long
    timestamp As Long
    forwarder As Long
    pDllName As Long
    pThunk As Long
End Type

Public Modules As New Collection  'of CImport
Public offset As CPEOffsets
Public errMessage As String

Function LoadImports(fpath As String) As Boolean
    On Error Resume Next
    
    Dim fOffsetIDT As Long
    Dim idt As IMAGE_IMPORT_DIRECTORY
    
    Set offset = New CPEOffsets
    Set Modules = New Collection
    
'    If InStr(fpath, "ntdll") > 0 Then
'        Stop
'    End If
    
    If Not offset.LoadFile(fpath) Then
        errMessage = offset.errMessage
        Exit Function
    End If
    
    If offset.RvaImportDirectory = 0 Then
        errMessage = "No Imports: "
        Exit Function
    End If
    
    fhandle = FreeFile
    Open fpath For Binary Access Read As fhandle
    
    fOffsetIDT = offset.RvaToOffset(offset.RvaImportDirectory)
    
    If fOffsetIDT < 1 Then
        errMessage = errMessage & vbCrLf & "Error getting file offset from rva? rva=" & offset.RvaImportDirectory & " offset=" & fOffsetIDT
        Exit Function
    End If
     
    fOffsetIDT = fOffsetIDT + 1 'cause 0 based file offset
    
    Get fhandle, fOffsetIDT, idt
    
    While idt.pFuncAry > 0
        LoadModuleImports idt
        fOffsetIDT = fOffsetIDT + Len(idt)
        'Debug.Print "offset=" & Hex(fOffsetIDT)
        Get fhandle, fOffsetIDT, idt
    Wend
    
    If Err.Number = 0 Then
        LoadImports = True
    Else
        errMessage = Err.Description
    End If
    
End Function

Private Sub LoadModuleImports(iid As IMAGE_IMPORT_DIRECTORY)
    'On Error GoTo hell
    
    Dim m  As New CImport

    With iid
        'Debug.Print "pFuncAry=" & Hex(.pFuncAry)
        m.pLookupTable = offset.RvaToOffset(.pFuncAry)
        m.DllName = ScanString(offset.RvaToOffset(.pDllName))
        m.FirstThunk = .pThunk
    End With
    
    Dim ImportLookup As Long, i As Long
    Dim isOrdial As Boolean, ImportFx As String
    
    i = m.pLookupTable + 1
    
    Get fhandle, i, ImportLookup
    isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
    ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)

        
    While ImportLookup > 0 Or isOrdial
    
        If isOrdial Then
            ImportFx = "@" & Hex(ImportLookup)
            'Debug.Print ImportFx
            m.functions.Add ImportFx
        Else
            ImportFx = ScanString(offset.RvaToOffset(ImportLookup), True)
            'Debug.Print ImportFx
            m.functions.Add ImportFx
        End If
        
        i = i + 4 'next record is 4 bytes away
        Get fhandle, i, ImportLookup
        isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
        ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)

    Wend
    
    Modules.Add m, CStr(m.DllName)

Exit Sub
hell: MsgBox Err.Description

End Sub

Private Function ScanString(offset As Long, Optional getHint As Boolean = False) As String
    Dim b As Byte, i As Long, ret As String
    Dim Hint As Integer '2 bytes
     
    i = offset + 1
    
    If getHint Then
        Get fhandle, i, Hint
        i = i + 2
    End If
    
    Get fhandle, i, b
     
    While b <> 0
        ret = ret & Chr(b)
        Get fhandle, , b
    Wend
    
    ScanString = ret
 
End Function


Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
