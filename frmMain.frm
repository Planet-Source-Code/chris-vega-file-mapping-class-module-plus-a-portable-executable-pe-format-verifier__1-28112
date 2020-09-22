VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check PE Format"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewPE 
      Caption         =   "Check PE"
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   390
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog dlgLoadEXE 
      Left            =   855
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.exe"
      DialogTitle     =   "Select PE-Executable"
      Filter          =   "Executable Files (*.exe)|*.exe|Dynamic Link Library (*.dll)|*.dll|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4890
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox PathFile 
      Height          =   285
      Left            =   750
      TabIndex        =   0
      Top             =   60
      Width           =   4125
   End
   Begin VB.Label lblFName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is a DEMO Application on Using my FileMapping VB Class Module
'   to Simulate the use of File Mapping APIs
'
'       Used in Visual C++ and Win32 Assembly, Why in Visual Basic, Why Not?
'
'
'   This code checks for a valid PE Executable Format
'
'   To understand this Application, you need to consult your nearest
'   PE Documentation.
'
'   Win32 Assembly Codes are included in Comments are 100% working on
'   TASM32 Compiler
'
'
'
'   Created by: Chris Vega [gwapo@models.com]
'               http://trider.8m.com
'
Private xPE As New clsFileMapping

Private Sub cmdBrowse_Click()
    dlgLoadEXE.ShowOpen
    If Not Err Then PathFile = dlgLoadEXE.FileName
End Sub

Private Sub cmdViewPE_Click()
    If xPE.OpenFile(Trim(PathFile)) Then
        ' Open the File with Read Access-Rights
        If xPE.MapFile Then
            ' Create File Mapping Object with Page-Read Access
            '   and Create a View of the Mapping Object
            '       all with Read-Only Access
            If xPE.OpenView Then
                '
                ' xPE Points to File Image Base
                ' ========================================================================
                '
                '       ; After File Mapping, set esi = ImageBase
                '
                '       ;   Note: ImageBase is the EntryPoint of a
                '       ;         Mapped File On View
                '
                '       xchg    eax, esi
                '       push    esi
                '       lodsw
                '       sub     ax,5a4dh
                '       jnz     @not_EXE
                '       pop     esi
                '
                With xPE
                    ' Load Word from the Image Base
                    MZ_Header = .lodsw
                    If Hex(MZ_Header) <> "5A4D" Then  ' ZM Marker Found?
                        '
                        ' @not_EXE:
                        '
                        MsgBox "Not an EXE File!", vbExclamation, "Chris ROQ!"
                        '
                        '   jmp     @close_File
                    Else
                        ' Go to PE Header RVA Address (ImageBase + 3Ch)
                        '
                        '   ; esi still holds the Image Base
                        '
                        '   push    esi                     ; Save Image Base
                        '   add     esi, 3ch                ; Point to lfaNew
                        '   lodsd                           ; Get the PE Header RVA
                        '   pop     esi                     ; Restore the Image Base
                        '   add     eax,esi                 ; Align PE Header RVA
                        '
                        '   ; eax now Holds the PE Header Virtual Address
                        '
                        .SetFilePointer &H3C, SetIncreaseFromCurrent
                        PE_Header = .lodsd
                        .SetFilePointer .GetFileEntryPoint + PE_Header, _
                                        SetReplaceCurrent
                        '
                        '   push    esi
                        '   xchg    eax,esi
                        '   lodsd
                        '   sub     ax,4550h
                        '   jnz     @not_PE
                        '   pop     esi
                        '
                        PE_Header = .lodsd
                        
                        If Hex(PE_Header) = "4550" Then   ' EP Marker Found?
                            '
                            ' ;  We got a Valid PE File,
                            ' ;    Lets go Check what type of PE is this!
                            '
                            '   mov     ax, word ptr [esi+18h]
                            '   sub     ax, 010bh               ; PE32?
                            '   jz      @found_PE32
                            '   jmp     @found_PE64
                            '
                            '   ; 010b = PE32
                            '   ; 020b = PE32+/PE64
                            '
                            .SetFilePointer &H18, SetIncreaseFromCurrent
                            OH_Header = .lodsd
                            
                            OH_Magic = Right(Hex(OH_Header), 4)
                            
                            If OH_Magic = "010B" Then _
                                MsgBox "This is a valid PE32 Executable File!", _
                                        vbInformation, _
                                        "Chris Vega [gwapo@models.com]" Else _
                                MsgBox "This is a valid PE32+/PE64 Executable File!", _
                                        vbInformation, _
                                        "Chris Vega [gwapo@models.com]"
                        Else
                            '
                            ' @not_PE:
                            MsgBox "Not a valid PE File!", _
                                   vbExclamation, _
                                   "Chris Vega [gwapo@models.com]"
                        End If
                    End If
                    .CloseView True     ' Close the View
                    .CloseMap           ' and the Mapping Object
                    .CloseFile          '   finally, Close the File
                End With
            End If
        Else
            MsgBox "Error Creating Mapping Object.", vbExclamation
        End If
        '
        ' @close_File:
        '
        xPE.CloseFile
    Else
        MsgBox "Error Opening the File", vbExclamation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Copyright 2001 by Chris Vega, No Rights Reserved, Use Without Permission!", _
           vbInformation, "Chris Vega [gwapo@models.com]"
End Sub
