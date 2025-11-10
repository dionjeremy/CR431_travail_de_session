@{
    RootModule        = 'studentToolKit.psm1'
    ModuleVersion     = '1.1.0'
    Author            = 'Jeremy Dion, Gabriel Robert'
    Description       = 'Devoir pour le cour CR431 module'
    PowerShellVersion = '7.5'
    FunctionsToExport = '*'
    CmdletsToExport   = @('*')
    AliasesToExport   = @()
    #Module a importé avant l'utilisation de ce module
    RequiredModules   = @(
        @{ModuleName='Microsoft.Graph.Calendar';ModuleVersion='2.31.0'}
    )
}