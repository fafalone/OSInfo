VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtText1 
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OSVer As clsOSInfo

Private Sub Form_Load()

    Set OSVer = New clsOSInfo

    Dim s As String
    With OSVer
        s = s & vbCrLf & "OS Name: " & .OSName
        s = s & vbCrLf & "Service Pack ver.: " & .SPVer
        s = s & vbCrLf & "Is Server? " & .IsServer
        s = s & vbCrLf & "Bitness: " & .Bitness
        s = s & vbCrLf & "Is Win x64: " & .IsWin64
        s = s & vbCrLf & "Is Win x32: " & .IsWin32
        s = s & vbCrLf & "Edition: " & .Edition
        s = s & vbCrLf & "Suite mask: " & .SuiteMask
        s = s & vbCrLf & "ProductType: " & .ProductType
        s = s & vbCrLf & "PlatformID: " & .PlatformID & " (" & .Platform & ")"
        s = s & vbCrLf & "Is Domain controller: " & .IsDomainController
        s = s & vbCrLf & "Is Embedded: " & .IsEmbedded
        s = s & vbCrLf & "OS - XP/Server 2003(R2)? " & .IsWindowsXP
        s = s & vbCrLf & "OS - XP or newer? " & .IsWindowsXPOrGreater
        s = s & vbCrLf & "OS - XP SP3 or newer? " & .IsWindowsXP_SP3OrGreater
        s = s & vbCrLf & "OS - Vista/Server 2008? " & .IsWindowsVista
        s = s & vbCrLf & "OS - Vista or newer? " & .IsWindowsVistaOrGreater
        s = s & vbCrLf & "OS - 7/Server 2008R2? " & .IsWindows7
        s = s & vbCrLf & "OS - 7 or newer? " & .IsWindows7OrGreater
        s = s & vbCrLf & "OS - 8/Server 2012? " & .IsWindows8
        s = s & vbCrLf & "OS - 8 or newer? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 8.1/Server 2012R2? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 8.1 or newer? " & .IsWindows8Point1OrGreater
        s = s & vbCrLf & "OS - 10/Server 2016? " & .IsWindows10
        s = s & vbCrLf & "OS - 10 or newer? " & .IsWindows10OrGreater
        s = s & vbCrLf & "OS - 11 or newer? " & .IsWindows11OrGreater
        s = s & vbCrLf & "Major: " & .Major
        s = s & vbCrLf & "Minor: " & .Minor
        s = s & vbCrLf & "Major + Minor:         " & .MajorMinor
        s = s & vbCrLf & "Major + Minor (NtDll): " & .MajorMinorNTDLL
        s = s & vbCrLf & "Build: " & .Build
        s = s & vbCrLf & "NT Dll Major.Minor.Rev: " & .NtDllVersion
        s = s & vbCrLf & "Revision: " & .Revision
        s = s & vbCrLf & "ReleaseId: " & .ReleaseId
        s = s & vbCrLf & "DisplayVersion: " & .DisplayVersion
        s = s & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
        s = s & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
        s = s & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
        s = s & vbCrLf & "ID of default locale: " & .LCID_UserDefault
        s = s & vbCrLf & "Process integrity level: " & .IntegrityLevel
        s = s & vbCrLf & "Elevated process? " & .IsElevated
        s = s & vbCrLf & "Is Local system context? " & .IsLocalSystemContext
        s = s & vbCrLf & "User name: " & .UserName
        s = s & vbCrLf & "User group: " & .UserType
        s = s & vbCrLf & "Is in Admin group? " & .IsAdminGroup
        s = s & vbCrLf & "Computer name: " & .ComputerName
        s = s & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
        s = s & vbCrLf & "Secure Boot supported? " & .SecureBootSupported & " (Enabled? " & .SecureBoot & ")"
        s = s & vbCrLf & "TestSigning: " & .TestSigning
        s = s & vbCrLf & "DebugMode: " & .DebugMode
        s = s & vbCrLf & "CodeIntegrity: " & .CodeIntegrity
        s = s & vbCrLf & "File System Case sensitive? " & .IsFileSystemCaseSensitive
        s = s & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
        s = s & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"
        s = s & vbCrLf & "Memory MiB (Free/Total): " & .MemoryFree & "/" & .MemoryTotal & " (Loaded: " & .MemoryLoad & "%)"
        s = s & vbCrLf & "CPU usage: " & .CpuUsage & "%"
        Debug.Print s
        txtText1.Text = s
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set OSVer = Nothing
End Sub
