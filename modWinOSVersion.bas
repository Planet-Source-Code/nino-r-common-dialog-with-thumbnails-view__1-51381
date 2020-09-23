Attribute VB_Name = "modWinOSVersion"
Option Explicit
'Visual Basic Helper Routines
'Handy Routines for Identifying the Windows Version - updated
      
'Posted:   Saturday August 14, 1999
'Updated:   utorak sijeèanj 27, 2004
      
'Applies to:   VB4-32, VB5, VB6
'Developed with:   VB6, Windows NT4
'OS restrictions:   None
'Author:   VBnet - Randy Birch

'dwPlatformId
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'os product type values
Public Const VER_NT_WORKSTATION = &H1
Public Const VER_NT_DOMAIN_CONTROLLER = &H2
Public Const VER_NT_SERVER = &H3

'product types
Public Const VER_SERVER_NT = &H80000000
Public Const VER_WORKSTATION_NT = &H40000000

Public Const VER_SUITE_SMALLBUSINESS = &H1
Public Const VER_SUITE_ENTERPRISE = &H2
Public Const VER_SUITE_BACKOFFICE = &H4
Public Const VER_SUITE_COMMUNICATIONS = &H8
Public Const VER_SUITE_TERMINAL = &H10
Public Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
Public Const VER_SUITE_EMBEDDEDNT = &H40
Public Const VER_SUITE_DATACENTER = &H80
Public Const VER_SUITE_SINGLEUSERTS = &H100
Public Const VER_SUITE_PERSONAL = &H200
Public Const VER_SUITE_BLADE = &H400

Public Const OSV_LENGTH As Long = 148
Public Const OSVEX_LENGTH As Long = 156

Public Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: string providing arbitrary additional information
End Type

Public Type OSVERSIONINFOEX
  OSVSize            As Long
  dwVerMajor        As Long
  dwVerMinor         As Long
  dwBuildNumber      As Long
  PlatformID         As Long
  szCSDVersion       As String * 128
  wServicePackMajor  As Integer
  wServicePackMinor  As Integer
  wSuiteMask         As Integer
  wProductType       As Byte
  wReserved          As Byte
End Type

'defined As Any to support OSVERSIONINFO and OSVERSIONINFOEX
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
'*** by nr
Public gnViewMenuPos As Integer

Public Function IsBladeServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows Server 2003 Web Edition is installed
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2003Server() Then
   
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
         IsBladeServer = (osv.wSuiteMask And VER_SUITE_BLADE)
      End If
   
   End If

End Function

Public Function IsDomainController() As Boolean
  
   Dim osv As OSVERSIONINFOEX
  'Returns True if the server is a domain
  'controller (Win 2000 or later)
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000Server() Then
   
      osv.OSVSize = Len(osv)
      
      If GetVersionEx(osv) = 1 Then
      
         IsDomainController = (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)
       
      End If
   
   End If

End Function

Public Function IsEnterpriseServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows NT 4.0 Enterprise Edition,
  'Windows 2000 Advanced Server, or Windows Server 2003
  'Enterprise Edition is installed.
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
   
      osv.OSVSize = Len(osv)
      
      If GetVersionEx(osv) = 1 Then
      
         IsEnterpriseServer = (osv.wProductType = VER_NT_SERVER) And _
                              (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
          
      End If
   
   End If

End Function

Public Function IsWin2000AdvancedServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows 2000 Advanced Server
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000Plus() Then
   
      osv.OSVSize = Len(osv)
      
      If GetVersionEx(osv) = 1 Then
      
         IsWin2000AdvancedServer = (osv.wProductType = VER_NT_SERVER) And _
                                   (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
                                   
      End If
   
   End If

End Function

Public Function IsWin2000Server() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows 2000 Server
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000() Then
   
      osv.OSVSize = Len(osv)
      
      If GetVersionEx(osv) = 1 Then
      
         IsWin2000Server = (osv.wProductType = VER_NT_SERVER)
         
      End If
   
   End If

End Function

Public Function IsSmallBusinessServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Microsoft Small Business Server is installed
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
   
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
         IsSmallBusinessServer = osv.wSuiteMask And VER_SUITE_SMALLBUSINESS
      End If
   
   End If

End Function

Public Function IsSmallBusinessRestrictedServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Microsoft Small Business Server
  'is installed with the restrictive client license
  'in force
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
   
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
         IsSmallBusinessRestrictedServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)
      End If
   
   End If

End Function

Public Function IsTerminalServer() As Boolean
  
   Dim osv As OSVERSIONINFOEX
  'Returns True if Terminal Services is installed
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
   
      osv.OSVSize = Len(osv)
      
      If GetVersionEx(osv) = 1 Then
         IsTerminalServer = osv.wSuiteMask And VER_SUITE_TERMINAL
      End If
   
   End If

End Function

Public Function IsWin95() As Boolean

  'returns True if running Win95
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin95 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                (osv.dwBuildNumber = 950)
                
   End If

End Function

Public Function IsWin95OSR2() As Boolean

  'returns True if running Win95
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin95OSR2 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                    (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                    (osv.dwBuildNumber = 1111)
 
   End If

End Function

Public Function IsWin98() As Boolean

  'returns True if running Win98
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin98 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 10) And _
                (osv.dwBuildNumber >= 2222)
                
   End If

End Function

Public Function IsWinME() As Boolean

  'returns True if running Windows ME
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinME = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 90) And _
                (osv.dwBuildNumber >= 3000)
     
   End If

End Function

Public Function IsWinNT4() As Boolean

  'returns True if running WinNT4
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinNT4 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                 (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                 (osv.dwBuildNumber >= 1381)
                 
   End If

End Function

Public Function IsWinNT4Plus() As Boolean

  'returns True if running Windows NT4 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                     (osv.dwVerMajor >= 4)
 
   End If

End Function

Public Function IsWinNT4Server() As Boolean

  'returns True if running Windows NT4 Server
   Dim osv As OSVERSIONINFOEX
      
   If IsWinNT4() Then
  
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
      
         IsWinNT4Server = (osv.wProductType And VER_NT_SERVER)
         
      End If

   End If

End Function

Public Function IsWinNT4Workstation() As Boolean

  'returns True if running Windows NT4 Workstation
   Dim osv As OSVERSIONINFOEX
      
   If IsWinNT4() Then
  
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
      
         IsWinNT4Workstation = (osv.wProductType And VER_NT_WORKSTATION)
         
      End If

   End If

End Function

Public Function IsWin2000() As Boolean

  'returns True if running Win2000 (NT5)
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin2000 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                  (osv.dwVerMajor = 5 And osv.dwVerMinor = 0) And _
                  (osv.dwBuildNumber >= 2195)
                  
   End If

End Function

Public Function IsWin2000Plus() As Boolean

  'returns True if running Windows 2000 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
  
   End If

End Function

Public Function IsWin2003Server() As Boolean

  'returns True if running Windows 2003 (.NET) Server
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWin2003Server = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                        (osv.dwVerMajor = 5 And osv.dwVerMinor = 2) And _
                        (osv.dwBuildNumber = 3790)

   End If

End Function

Public Function IsWin2000Workstation() As Boolean

  'returns True if running Windows NT4 Workstation
   Dim osv As OSVERSIONINFOEX
      
   If IsWin2000() Then
  
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
      
         IsWin2000Workstation = (osv.wProductType And VER_NT_WORKSTATION)
         
      End If

   End If

End Function

Public Function IsWinXP() As Boolean

  'returns True if running WinXP (NT5.1)
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinXP = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                (osv.dwVerMajor = 5 And osv.dwVerMinor = 1) And _
                (osv.dwBuildNumber >= 2600)

   End If

End Function

Public Function IsWinXPPlus() As Boolean

  'returns True if running WinXP (NT5.1) or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinXPPlus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1)

   End If

End Function

Public Function IsWinXPHomeEdition() As Boolean

  'returns True if running WinXP Home Edition (NT5.1)
   Dim osv As OSVERSIONINFOEX
      
   If IsWinXP() Then
  
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
      
         IsWinXPHomeEdition = (osv.wSuiteMask And VER_SUITE_PERSONAL)
         
      End If

   End If

End Function

Public Function IsWinXPProEdition() As Boolean

  'returns True if running WinXP Pro
   Dim osv As OSVERSIONINFOEX
      
   If IsWinXP() Then
  
      osv.OSVSize = Len(osv)
   
      If GetVersionEx(osv) = 1 Then
      
         IsWinXPProEdition = Not (osv.wSuiteMask And VER_SUITE_PERSONAL)
      
      End If

   End If

End Function

'--end block--'

Public Sub FindWindowsVersion()

    Dim bOSVersion As Boolean, sThisWinOS As String
   
    bOSVersion = IsWin95()
    If (bOSVersion) Then
        sThisWinOS = " Win95"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 4
        Exit Sub
    End If
    bOSVersion = IsWin95OSR2()
    If (bOSVersion) Then
        sThisWinOS = " Win95OSR2"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 4
        Exit Sub
    End If
    bOSVersion = IsWin98()
    If (bOSVersion) Then
        sThisWinOS = " Win98"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 4
        Exit Sub
    End If
    bOSVersion = IsWinME()
    If (bOSVersion) Then
        sThisWinOS = " WinME"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 4
        Exit Sub
    End If

    bOSVersion = IsWin2000()
    If (bOSVersion) Then
        sThisWinOS = " Win2000"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 4
        Exit Sub
    End If
    bOSVersion = IsWinXP()
    If (bOSVersion) Then
        sThisWinOS = " WinXP"
        frmMainT.txtOSVersion = sThisWinOS
        gnViewMenuPos = 0
        Exit Sub
    End If

'"Win95", "Win95 OSR2", "Win98", "WinME"
'" WinNT4", "WinNT4 Plus", "WinNT4 Server", "WinNT4 Workstation"
'"Win2000", "Win2000 Plus", "Win2000 Server"
'"Win2000 Advanced Server", "Win2000 Workstation"
'"Win2003 Server", "WinXP", "WinXP Plus"
'"WinXP Home", "WinXP Pro"
'"BackOffice Server","Blade Server"
'"Domain Controller","Enterprise Server"
'"Small Business Server", "Small Business Restricted Server"
'"Terminal Server"
   
'*** functions
'IsWin95(), IsWin95OSR2(), IsWin98(), IsWinME()
'IsWinNT4(), IsWinNT4Plus(), IsWinNT4Server(), IsWinNT4Workstation()
'IsWin2000(), IsWin2000Plus(), IsWin2000Server()
'IsWin2000AdvancedServer, IsWin2000Workstation()
'IsWin2003Server()
'IsWinXP(), IsWinXPPlus()
'IsWinXPHomeEdition(), IsWinXPProEdition()
'IsBackOfficeServer(), IsBladeServer()
'IsDomainController(), IsEnterpriseServer()
'IsSmallBusinessServer(), IsSmallBusinessRestrictedServer()
'IsTerminalServer()


End Sub



