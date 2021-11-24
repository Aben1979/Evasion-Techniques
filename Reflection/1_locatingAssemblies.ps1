# Search preloaded assemblies in the PowerShell process for static functions that have the keyword "unsafe" to run with C#
$Assemblies = [AppDomain]::CurrentDomain.GetAssemblies()
$Assemblies |
    ForEach-Object {
        $_.GetTypes()|
            ForEach-Object {
                $_ | Get-Member -Static| Where-Object {
                    $_.TypeName.Contains('Unsafe')
                    }
            } 2> $null
    }
# Locate GetModuleHandle and GetProcAddress in the output (In the Microsoft.Win32.UnsafeNativeMethods class)

# Modified query to print the current assembly location through the Location property and then inside the nested ForEach-Object loop make the TypeName match Microsoft.Win32.UnsafeNativeMethods instead of listing all methods with the static keyword:

$Assemblies = [AppDomain]::CurrentDomain.GetAssemblies()
$Assemblies |
    ForEach-Object {
        $_.Location
        $_.GetTypes()|
            ForEach-Object {
                $_ | Get-Member -Static| Where-Object {
                    $_.TypeName.Equals('Microsoft.Win32.UnsafeNativeMethods')
                }
            } 2> $null
    }
# The assembly is C:\Windows\Microsoft.Net\assembly\GAC_MSIL\System\v4.0_4.0.0.0__b77a5c561934e089\System.dll

# Obtain a reference to the System.dll assembly using the GetType147 method:
$systemdll = ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.GlobalAssemblyCache -And $_.Location.Split('\\')[-1].Equals('System.dll') })
$unsafeObj = $systemdll.GetType('Microsoft.Win32.UnsafeNativeMethods')
$GetModuleHandle = $unsafeObj.GetMethod('GetModuleHandle')
$GetModuleHandle.Invoke($null, @("user32.dll")) # Resolve the user32.dll assembly location
1981612032 # <<<<<<<< The output is the assembly address

# Use refelction again through GetMethod to locate GetProcAddress and use it to resolve the MessageBoxA function address
$systemdll = ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {
    $_.GlobalAssemblyCache -And $_.Location.Split('\\')[-1].Equals('System.dll') })
$unsafeObj = $systemdll.GetType('Microsoft.Win32.UnsafeNativeMethods')
$GetModuleHandle = $unsafeObj.GetMethod('GetModuleHandle')
$user32 = $GetModuleHandle.Invoke($null, @("user32.dll"))
$tmp=@()
$unsafeObj.GetMethods() | ForEach-Object {If($_.Name -eq "GetProcAddress") {$tmp+=$_}}
$GetProcAddress = $tmp[0]
$GetProcAddress.Invoke($null, @($user32, "MessageBoxA"))

# The method resolveWin32API.ps1 in the same directory performs the above steps to resolve an Win32 API without using Add-Type