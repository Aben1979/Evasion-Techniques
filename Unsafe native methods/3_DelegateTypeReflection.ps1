# create a new assembly object through the AssemblyName class and assign it a name like ReflectedDelegate
$MyAssembly = New-Object System.Reflection.AssemblyName('ReflectedDelegate')

# Configure the access mode to be executable and not saved to disk:
$Domain = [AppDomain]::CurrentDomain
$MyAssemblyBuilder = $Domain.DefineDynamicAssembly($MyAssembly, [System.Reflection.Emit.AssemblyBuilderAccess]::Run)

# Create a module through the DefineDynamicModule method. supply a custom name for the module and tell it not to include symbol information:
$MyModuleBuilder = $MyAssemblyBuilder.DefineDynamicModule('InMemoryModule', $false)

# Create a custom type:
$MyTypeBuilder = $MyModuleBuilder.DefineType('MyDelegateType', 'Class, Public, Sealed, AnsiClass, AutoClass', [System.MulticastDelegate])

# Put the function prototype inside the type and let it become the custom delegate type:
$MyConstructorBuilder = $MyTypeBuilder.DefineConstructor('RTSpecialName, HideBySig, Public', [System.Reflection.CallingConventions]::Standard, @([IntPtr], [String], [String], [int]))

# Set implementation flags before calling the constructor:
$MyConstructorBuilder.SetImplementationFlags('Runtime, Managed')

# Specify settings for the "Invoke" method:
$MyMethodBuilder = $MyTypeBuilder.DefineMethod('Invoke', 'Public, HideBySig, NewSlot, Virtual', [int], @([IntPtr], [String], [String], [int]))

# Set implementation flags
$MyDelegateType = $MyTypeBuilder.CreateType()