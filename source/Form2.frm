VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Entry"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1020
      TabIndex        =   11
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1020
      TabIndex        =   10
      Top             =   1260
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   9
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1020
      TabIndex        =   8
      Top             =   660
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   7
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   6
      Top             =   60
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "HookMod"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "HookProc"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "1st Inst"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "MemAdr"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "IAT"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Sub ShowItem(li As ListItem)
    
    On Error Resume Next
    Text1(0) = li.Text
    Text1(1) = li.SubItems(1)
    Text1(2) = li.SubItems(2)
    Text1(3) = li.SubItems(3)
    Text1(4) = li.SubItems(4)
    Text1(5) = li.SubItems(5)
    
    Me.Show 1
    If Err.Number > 0 Then Me.Visible = True
    
End Sub

