VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Hook Explorer  (Detects IAT and basic Detours style hooks for bound & dynamic loaded imgs)"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   60
      TabIndex        =   14
      Top             =   5580
      Width           =   9495
      Begin VB.TextBox txtError 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   9375
      End
      Begin VB.CommandButton cmdreload 
         Caption         =   "Reload"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Message Log"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "IgnoreList"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.OptionButton optDisplay 
      Caption         =   "Standard"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   2700
      Width           =   1035
   End
   Begin VB.OptionButton optDisplay 
      Caption         =   "Show All entries"
      Height          =   255
      Index           =   3
      Left            =   7980
      TabIndex        =   12
      Top             =   2700
      Width           =   1455
   End
   Begin VB.OptionButton optDisplay 
      Caption         =   "Hide Hooks within same module"
      Height          =   255
      Index           =   2
      Left            =   5340
      TabIndex        =   11
      Top             =   2700
      Width           =   2775
   End
   Begin VB.OptionButton optDisplay 
      Caption         =   "Use Ignore List"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   2700
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.CheckBox chkScanExports 
      Caption         =   "Scan all exports"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2700
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvBound 
      Height          =   2355
      Left            =   3780
      TabIndex        =   3
      Top             =   240
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BaseAdr"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MaxAdr"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hooks"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "pid"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "process"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "user"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvImports 
      Height          =   2475
      Left            =   60
      TabIndex        =   5
      Top             =   3000
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IAT Address"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "1st Instruction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HookProc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "HookMod"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2340
      TabIndex        =   9
      Top             =   2700
      Width           =   135
   End
   Begin VB.Label mnuRefresh 
      BackColor       =   &H8000000B&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Functions"
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Processes "
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblImage 
      Caption         =   "Dlls"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyItem 
         Caption         =   "Copy Line"
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "View Line"
      End
      Begin VB.Menu mnuCopyList 
         Caption         =   "Export Table"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyDllList 
         Caption         =   "Copy Dll List"
      End
      Begin VB.Menu mnuTabCopyList 
         Caption         =   "Export Results to Tab List"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         disassembler functionality provided by olly.dll which
'         is a modified version of the OllyDbg GPL source from
'         Oleh Yuschuk Copyright (C) 2001 - http://ollydbg.de
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



Private Type t_Disasm
  ip As Long
  dump As String * 256
  result As String * 256
  unused(1 To 308) As Byte
End Type

Public Enum displayOpt
    doStandard = 0
    doShowAll = 3    'std/all very close only diff is all shows all imports/exports dll
    doIgnoreList = 1
    doHideSelf = 2
End Enum
    
Private Declare Function Disasm Lib "olly.dll" (ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, Disasm As t_Disasm, Optional disasmMode As Long = 4) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessLong Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessBuffer Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const PROCESS_VM_READ = (&H10)

Dim proc As New CProcessInfo
Dim imports As New CLoadImports
Dim exports As New CLoadExports
Dim pe As New CPEOffsets

Dim Modules As New Collection
Dim Containers As New Collection
Dim AnalyzedFx As New Collection
Public Symbols As New Collection

Dim liProc As ListItem
Dim liEntry As ListItem
Dim h As Long                  'global openprocess handle
Public DisplayOption As displayOpt
Public IgnoreList As New Collection  'ignore these dlls



Private Sub cmdEdit_Click()
    Dim lst As String
    
    lst = App.path & IIf(IsIde, "\..\", "") & "\IgnoreList.txt"
    
    If Not FileExists(lst) Then
        MsgBox "Could not find ignore list?" & lst, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & lst & """", vbNormalFocus
    
    If Err.Number > 0 Then
        MsgBox "Could not spawn notepad"
    End If
    
End Sub

Private Sub cmdreload_Click()
    Dim tmp() As String
    Dim x
    Dim lst As String
    Dim c As CContainer
    
    Set IgnoreList = New Collection
    
    lst = App.path & IIf(IsIde, "\..\", "") & "\IgnoreList.txt"
    
    If Not FileExists(lst) Then
        MsgBox "Could not find ignore list?" & lst, vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    
    tmp = Split(ReadFile(lst), vbCrLf)
    For Each x In tmp
        If Left(x, 1) <> "#" And Left(x, 1) <> ";" And Left(x, 1) <> "/" Then
            x = LCase(x)
            IgnoreList.Add x, x
        End If
    Next
    
    If lv.ListItems.Count = 0 Then Exit Sub 'from form_load event
    
    For Each c In Containers
        c.ReapplyFilters
    Next
    
    DisplayModules True
    
End Sub

Private Sub Form_Load()

    Dim d As New CProcess
    Dim li As ListItem
    Dim p As New Collection
    Dim tmp() As String
    Dim x
    
    On Error Resume Next
     
    lv.ListItems.Clear
    lvBound.ListItems.Clear
    lvImports.ListItems.Clear
    
    SizeLV lv
    SizeLV lvBound
    SizeLV lvImports
     
    cmdreload_Click 'load ignore list
    chkScanExports.value = GetSetting("PHE", "PHE", "chkScanExports", 1)
    DisplayOption = GetSetting("PHE", "PHE", "optDisplay", 1)
    optDisplay(DisplayOption).value = 1
    
    Set p = proc.GetRunningProcesses
    
    For Each d In p
        If d.pid > 0 And d.pid <> 8 Then
            Set li = lv.ListItems.Add(, , d.pid)
            li.SubItems(1) = d.path
            li.SubItems(2) = d.User
            li.Tag = d.pid
        End If
    Next
    
    If Not proc.GetSeDebug Then
        MsgBox "Could not get SEDebug Privledge" & vbCrLf & vbCrLf & _
                "You should run this as administrator " & vbCrLf & _
                "or power user if possible", vbInformation
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height < 8100 Then Me.Height = 8100
    If Me.Width < 9660 Then Me.Width = 9660
    
    lvBound.Width = Me.Width - lvBound.Left - 200
    lvImports.Width = Me.Width - lvImports.Left - 200
    Frame1.Width = Me.Width
    txtError.Width = lvImports.Width
    SizeLV lvBound
    SizeLV lvImports
    
    Frame1.Top = Me.Height - Frame1.Height - 300
    lvImports.Height = Me.Height - Frame1.Height - lvImports.Top - 400
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "PHE", "PHE", "chkScanExports", chkScanExports.value
End Sub

Private Sub lv_Click()

    If liProc Is Nothing Then Exit Sub
    
    Dim m As CModule
    Dim x As CContainer

    Set Symbols = New Collection    'key: memadr , value : dllname!funcname
    Set Containers = New Collection 'holds CContainer objects: key dllname
    Set AnalyzedFx = New Collection 'stores unique dll:function list
    Set Modules = proc.GetProcessModules(CLng(liProc.Text))
    
    txtError = Empty
    pb.value = 0
    pb.Visible = True
    pb.Max = Modules.Count + 1
    lvBound.ListItems.Clear
    lvImports.ListItems.Clear
    lvImports.ListItems.Clear
    lblImage = "Scanning.."
    
    h = OpenProcess(PROCESS_VM_READ, 0, CLng(liProc.Text))
    
    LogMsg "Scanning for hooks in:" & liProc.SubItems(1) & " - " & (Modules.Count - 1) & " dlls in this process"
    LogMsg String(75, "-")

    For Each m In Modules
        Set x = GetContainer(FileNameFromPath(m.path), m)
        AnalyzeIAT x.Module.path, x
        incpb
    Next
    
    DisplayModules
    
    CloseHandle h
    pb.Visible = False
    lblImage = "Hooked Dlls for " & liProc.SubItems(1)
        
End Sub

Sub DisplayModules(Optional clearList As Boolean = False)
    
    Dim m As CModule
    Dim li As ListItem
    Dim x As CContainer
    Dim i As Long
    Dim tmp As Long
    Dim c As Collection
    
    If Modules.Count = 0 Then Exit Sub
    
    If clearList Then
        txtError = Empty
        lvBound.ListItems.Clear
        lvImports.ListItems.Clear
        For Each x In Containers
            Set li = lvBound.ListItems.Add(, , Hex(x.Module.base))
            Set x.li = li
            Set li.Tag = x
            li.SubItems(1) = x.size
            li.SubItems(3) = x.DllName
        Next
        LogMsg "Scanning for hooks in:" & liProc.SubItems(1) & " - " & (Modules.Count - 1) & " dlls in this process"
        LogMsg String(75, "-")
    End If
        
    'prune list to show only hooked dlls
    'update our listview with final stats
    For i = lvBound.ListItems.Count To 1 Step -1
        Set li = lvBound.ListItems(i)
        Set x = li.Tag
           
        Select Case DisplayOption
            Case doShowAll, doStandard:    tmp = x.AllHookedEntries.Count
            Case doIgnoreList:             tmp = x.FilteredHookedEntries.Count
            Case doHideSelf:               tmp = x.RealHookedEntries.Count
        End Select
        
        If tmp = 0 Then
        
            LogMsg "No hooks - " & pad(x.DllName, 12) & _
                    pad("(0x" & Hex(x.Module.base) & ")", 14) & _
                   "Exports Scanned: " & IIf(x.ExportsScanned, pad(x.ExportsTotal), "False ") & _
                   IIf(x.Module.Rebased, "Rebased", "")
                   
            lvBound.ListItems.Remove li.index
            
        Else
            Select Case DisplayOption
                Case doShowAll:    li.SubItems(2) = x.AllHookedEntries.Count & " / " & x.Entries.Count
                Case doStandard:   li.SubItems(2) = x.RealHookedEntries.Count
                Case doIgnoreList: li.SubItems(2) = x.FilteredHookedEntries.Count
                Case doHideSelf:   li.SubItems(2) = x.RealHookedEntries.Count
            End Select
        End If
    Next
    
    'log final stats to message window
    For Each li In lvBound.ListItems
        Set x = li.Tag
        
        Select Case DisplayOption
            Case doShowAll, doStandard:   tmp = x.AllHookedEntries.Count
            Case doIgnoreList:            tmp = x.FilteredHookedEntries.Count
            Case doHideSelf:              tmp = x.RealHookedEntries.Count
        End Select
        
        LogMsg pad(tmp, 3) & " hooks in " & _
               pad(x.DllName, 12) & _
               pad("(0x" & Hex(x.Module.base) & ")", 14) & _
               "Exports Scanned: " & IIf(x.ExportsScanned, pad(x.ExportsTotal), "False ") & _
               IIf(x.Module.Rebased, "Rebased", "")
               
    Next
    
    
End Sub


Sub AnalyzeIAT(fpath As String, x As CContainer)
    
    Dim mtmp As CModule
    Dim xtmp As CContainer
    Dim c As CImport

    On Error GoTo hell
    
    If Not pe.LoadFile(fpath) Then
        LogMsg "***** ERROR: could not load pefile: " & fpath
        Exit Sub
    End If
    
    If pe.RvaImportDirectory = 0 Then
        LogMsg "No imports for - " & fpath
    Else
        If Not imports.LoadImports(fpath) Then
            LogMsg "***** ERROR: load import failed: " & fpath
        Else
            'this modules IAT entries are held in a collection
            'with each imported dll as a seperate CImport entry
            'So here we scan through them all for our target module
            If Not x.ImportsScanned Then
                x.ImportsScanned = True
                For Each c In imports.Modules
                    Set mtmp = GetModuleFor(c.DllName)
                    AnalyzeImports x, mtmp, c
                    DoEvents
                Next
            End If
        End If
    End If
    
    'If InStr(fpath, "ntdll.dll") > 0 Then Stop
    
    'if dll was dynamically loaded this might be only chance to scan
    If chkScanExports.value = 1 And pe.RvaExportDirectory > 0 Then
        If Not x.ExportsScanned Then
            If Not exports.LoadExports(fpath) Then
                LogMsg "***** ERROR: Failed to load Exports: " & fpath
            Else
                x.ExportsTotal = exports.functions.Count
                AnalyzeExports x
            End If
        End If
    End If
    
Exit Sub
hell:
        Debug.Print "Hell in AnalyzeIAT: " & x.DllName
        
End Sub

Sub AnalyzeImports(x As CContainer, m As CModule, c As CImport)
    
    Dim fx, d As String
    Dim FirstThunk As Long
    Dim ptr As Long
    Dim v As Long
    Dim isHooked As Boolean
    Dim HookAdr As Long, HookMod As String
        
    On Error GoTo hell
      
    FirstThunk = c.FirstThunk
    
    For Each fx In c.functions
        ptr = x.Module.base + FirstThunk
        ReadProcessLong h, ptr, v, 4, 0 'get adr from IAT
        
        HookAdr = v
        isHooked = m.AddressInRange(v)
        HookMod = GetModuleForAddress(v)
         
        ParseHookVals v, d, HookAdr, HookMod, isHooked
        x.AddEntry ptr, CStr(fx), v, d, c.DllName, isHooked, HookAdr, HookMod
        
        DoEvents
        FirstThunk = FirstThunk + 4 'goto next IAT entry mem addr
    Next
    
Exit Sub
hell:
        Debug.Print "hell in AnalyzeImports: " & c.DllName
    
End Sub


Sub AnalyzeExports(x As CContainer)
    
    Dim fx As CExport
    Dim d As String
    Dim FirstThunk As Long
    Dim ptr As Long
    Dim v As Long
    Dim bytes() As Byte
    Dim isHooked As Boolean
    Dim HookAdr As Long, HookMod As String
    Dim key As String
    Dim name As String
    Dim f As Long
    
    On Error GoTo hell
    
    'If x.ExportsScanned Then Stop

    x.ExportsScanned = True
    
    For Each fx In exports.functions
        
        If Len(fx.FunctionName) > 0 Then name = fx.FunctionName _
        Else name = "@" & fx.FunctionOrdial
        
        key = x.DllName & ":" & LCase(name)
        
        If KeyExistsInCollection(AnalyzedFx, key) Then GoTo nextExport _
        Else AnalyzedFx.Add key, key
        
        isHooked = False
        HookAdr = 0
        HookMod = Empty
        ReDim bytes(16)
        
        ptr = x.Module.base + fx.FunctionAddress
            
        If Not KeyExistsInCollection(Symbols, "adr:" & Hex(ptr)) Then
            Symbols.Add Replace(x.DllName, ".dll", Empty) & "." & name, "adr:" & Hex(ptr)
        End If
        
        ParseHookVals ptr, d, HookAdr, HookMod, isHooked
        x.AddEntry -1, name, HookAdr, d, x.DllName, isHooked, HookAdr, HookMod
        
        
nextExport:
        DoEvents
    Next
    
Exit Sub
hell:
    Debug.Print "Hell in AnalyzeExports: " & x.DllName & " Desc:" & Err.Description
End Sub

'this fx either gets an existing CContainer for a given
'dllname or else it creates and configures a new one.
Function GetContainer(ByVal DllName As String, m As CModule) As CContainer
    Dim dll As String
    Dim x As CContainer
    
    If InStr(1, DllName, "\") > 0 Then DllName = FileNameFromPath(DllName)
    dll = LCase(DllName)
    
    On Error Resume Next
    
    If objKeyExistsInCollection(Containers, dll) Then
        Set x = Containers(dll)
    Else
        Set x = New CContainer
        Set x.li = lvBound.ListItems.Add(, , Hex(m.base))
        Set x.Module = m
        Set x.li.Tag = x
        x.size = Hex(m.base + m.size)
        x.li.SubItems(1) = x.size
        x.li.SubItems(3) = DllName
        x.DllName = DllName
        Containers.Add x, dll
    End If

    Set GetContainer = x
        
End Function

'** all args except ptr are out vals
'gets 1st inst disasm from mem addr and parses it looking for hooks
Sub ParseHookVals(ByVal ptr As Long, d As String, HookAdr As Long, HookMod As String, isHooked As Boolean)
    On Error Resume Next
    Dim bytes(16) As Byte
    
    If ReadProcessBuffer(h, ptr, bytes(0), 15, 0) <> 0 Then
        d = DisasmBytes(bytes, ptr)
    End If
 
    'what is the first instruction of the target function, does it jmp off right away
    'i know more elaborate hooks wont be detected...some fx like some in msvcrt that
    'immediatly call a local function will show as local hooks...could hide but..
    If InStr(1, d, "jmp", vbTextCompare) > 0 Or InStr(1, d, "call", vbTextCompare) Then
        isHooked = True
        If InStr(d, "[") > 0 Then 'get actual code address from pointer
            ptr = CLng("&h" & Replace(Trim(Mid(d, InStrRev(Trim(d), "[") + 1)), "]", Empty))
            If ptr > 0 Then 'should change to isbadreadptr(ptr) probably
                ReadProcessLong h, ptr, HookAdr, 4, 0
            Else
                HookAdr = -1
            End If
        Else
            If InStr(d, ":") > 0 Then 'not supported (probably bad disasm anyway)
                HookAdr = -1
            Else
                HookAdr = CLng("&h" & Trim(Mid(d, InStrRev(Trim(d), " "))))
            End If
        End If
        HookMod = GetModuleForAddress(HookAdr)
    End If
        
    
End Sub

Function GetModuleFor(modName) As CModule
    
    Dim d As CModule
    
    For Each d In Modules
        If InStr(1, d.path, modName, vbTextCompare) > 0 Then
            Set GetModuleFor = d
            Exit Function
        End If
    Next
    
End Function

Function GetModuleForAddress(v As Long) As String
    
    Dim d As CModule
    
    If v = 0 Then Exit Function
    
    For Each d In Modules
        If v >= d.base And v <= (d.base + d.size) Then
            GetModuleForAddress = d.path
            Exit Function
        End If
    Next
    
    GetModuleForAddress = "Unknown"
    
End Function

'click on a dll name to fill out hook list
Private Sub lvBound_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim c As CContainer
    Set c = Item.Tag
    c.FillOutListView
End Sub


'----------------------------------------------------------------
'generic library functions below
'----------------------------------------------------------------

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Function DisasmBytes(b() As Byte, va As Long) As String
    Dim da As t_Disasm
    On Error Resume Next
    Disasm b(0), UBound(b) + 1, va, da
    DisasmBytes = Mid(da.result, 1, InStr(da.result, Chr(0)) - 1)
End Function

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Function objKeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    objKeyExistsInCollection = True
 Exit Function
nope: objKeyExistsInCollection = False
End Function

Sub LogMsg(msg)
    txtError.SelStart = Len(txtError)
    txtError.SelText = msg & vbCrLf
End Sub



Function FileNameFromPath(fullpath) As String
    On Error Resume Next
    Dim tmp() As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function


Function pad(v, Optional l As Long = 6) As String
        Dim tmp As String
        Dim x As Long
        
        tmp = CStr(v)
        x = l - Len(tmp)
        If x > 0 Then tmp = tmp & Space(x)
        pad = tmp
        
End Function

Private Sub incpb()
    On Error Resume Next
    pb.value = pb.value + 1
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liProc = Item
End Sub


Private Sub Label4_Click()
    
    MsgBox "This option will make it scan the entire export table" & vbCrLf & _
            "for each dll found in memory looking for detours style" & vbCrLf & _
            "hooks. " & vbCrLf & _
            "" & vbCrLf & _
            "For dynamically loaded dlls, this would be the only" & vbCrLf & _
            "chance we get to perform some kind of check on " & vbCrLf & _
            "them. ", vbInformation
    
End Sub

 

Private Sub lvBound_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup2
End Sub

Private Sub lvImports_DblClick()
    If liEntry Is Nothing Then Exit Sub
    Form2.ShowItem liEntry
End Sub

Private Sub lvImports_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liEntry = Item
End Sub

Private Sub lvImports_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyDllList_Click()
    
    On Error Resume Next
    
    Dim f As String
    Dim li As ListItem
    Dim tmp()
    
    If lvBound.ListItems.Count = 0 Then Exit Sub
    
    For Each li In lvBound.ListItems
        With li
            push tmp, "DllBase : " & .Text
            push tmp, "MaxAddr : " & .SubItems(1)
            push tmp, "Hooks   : " & .SubItems(2)
            push tmp, "DllName : " & .SubItems(3)
            push tmp, ""
        End With
    Next
    
    f = WriteTmpFile(Join(tmp, vbCrLf))

    If Len(f) = 0 Then
        MsgBox "Could not generate tmp file", vbInformation
        Exit Sub
    End If
    
    Shell "notepad """ & f & """", vbNormalFocus
    
    Sleep 800
    DoEvents
    Kill f
    
End Sub

Private Sub mnuCopyItem_Click()
    
    If liEntry Is Nothing Then Exit Sub
    
    Dim tmp()
    
    With liEntry
        push tmp, "IAT Addr: " & .Text
        push tmp, "HookAddr: " & .SubItems(1)
        push tmp, "1st Inst: " & .SubItems(2)
        push tmp, "Name    : " & .SubItems(3)
        push tmp, "HookProc: " & .SubItems(4)
        push tmp, "HookMod : " & .SubItems(5)
        Clipboard.Clear
        Clipboard.SetText Join(tmp, vbCrLf)
    End With
    
    MsgBox "Data copied to clipboard"
     
End Sub

Private Sub mnuCopyList_Click()
    On Error Resume Next
    
    Dim f As String
    Dim li As ListItem
    Dim tmp()
 
    If lvImports.ListItems.Count = 0 Then Exit Sub
    
    For Each li In lvImports.ListItems
        With li
            push tmp, "IAT Addr: " & .Text
            push tmp, "HookAddr: " & .SubItems(1)
            push tmp, "1st Inst: " & .SubItems(2)
            push tmp, "Name    : " & .SubItems(3)
            push tmp, "HookProc: " & .SubItems(4)
            push tmp, "HookMod : " & .SubItems(5)
            push tmp, "" & String(50, "-") & vbCrLf
        End With
    Next
    
    f = WriteTmpFile(Join(tmp, vbCrLf))

    If Len(f) = 0 Then
        MsgBox "Could not generate tmp file", vbInformation
        Exit Sub
    End If
    
    Shell "notepad """ & f & """", vbNormalFocus
    
    Sleep 800
    DoEvents
    Kill f
    
    
End Sub

Private Sub mnuTabCopyList_Click()
    On Error Resume Next
    Dim f As String
    Dim fHandle As Long
    Dim x As CContainer
    
    Const header = "IAT Address\tValue\tName\t1stInst\tHookProc\tHookMod"
    
    If lvBound.ListItems.Count = 0 Then
        MsgBox "No results to export", vbInformation
        Exit Sub
    End If
    
    f = WriteTmpFile("") 'get tmp file name
    fHandle = FreeFile
    
    Open f For Output As #fHandle
    Print #fHandle, Replace(header, "\t", vbTab)
                    
    For Each x In Containers
        x.DumpSelectedToHandle fHandle
    Next
    
    Close #fHandle
    
    MsgBox "Make sure to save results", vbInformation
        
    Shell "notepad """ & f & """", vbNormalFocus
    
    Sleep 800
    DoEvents
    Kill f
    
End Sub

Private Sub mnuRefresh_Click()
    Form_Load
End Sub

Sub SizeLV(lv As ListView)
    Dim c As Long
    With lv
        c = .ColumnHeaders.Count
        .ColumnHeaders(c).Width = .Width - lv.ColumnHeaders(c).Left - 300
    End With
End Sub

Private Sub cmdRefresh_Click()
    Form_Load
End Sub

Private Sub mnuSearchList_Click()
    MsgBox "ToDo"
End Sub



Private Sub mnuViewItem_Click()
    If liEntry Is Nothing Then Exit Sub
    Form2.ShowItem liEntry
End Sub

Private Sub optDisplay_Click(index As Integer)
    SaveSetting "PHE", "PHE", "optDisplay", index
    DisplayOption = index
    If Modules.Count > 0 Then DisplayModules True
End Sub

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
  Open filename For Binary As #f        ' Open file.(can be text or image)
  temp = Input(FileLen(filename), #f)   ' Get entire Files data
  Close #f
  ReadFile = temp
End Function



Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 \ 0
    IsIde = False
Exit Function
hell: IsIde = True
End Function

'Private Sub mnuSHowDlls_Click()
'    Dim c As Collection
'    Dim m As CModule
'    Dim tmp As String
'
'    If liProc Is Nothing Then Exit Sub
'
'    Set c = proc.GetProcessModules(liProc.Text)
'
'    For Each m In c
'        tmp = tmp & m.path & vbCrLf
'    Next
'
'    On Error Resume Next
'    Dim f As String
'    f = WriteTmpFile(tmp)
'    Shell "notepad """ & f & """", vbNormalFocus
'    DoEvents
'    Kill f
'
'End Sub

Function WriteTmpFile(it) As String
    Dim f As Long
    Dim tmp As String
    On Error GoTo hell

    tmp = Environ("TEMP")
    If Len(tmp) = 0 Then tmp = Environ("TMP")
    If Len(tmp) = 0 Then tmp = "C:\"
    If Right(tmp, 1) <> "\" Then tmp = tmp & "\"
    tmp = tmp & GetTickCount & ".txt"

    f = FreeFile
    Open tmp For Output As #f
    Print #f, it
    Close f

    WriteTmpFile = tmp

hell:
End Function
 

Private Sub txtError_DblClick()
    
    On Error Resume Next
    Dim f As String
    
    f = WriteTmpFile(txtError.Text)
    
    If Len(f) = 0 Then
        MsgBox "Could not generate tmp file", vbInformation
        Exit Sub
    End If
    
    Shell "notepad """ & f & """", vbNormalFocus
    
    Sleep 800
    DoEvents
    Kill f

End Sub




Function GetAllElements(lv As ListView) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem
    
    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next
    
    push ret, tmp
    
    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.Count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next
    
    GetAllElements = Join(ret, vbCrLf)
    
End Function


