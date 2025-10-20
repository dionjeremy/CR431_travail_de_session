@{
    RootModule        = 'devoir.psm1'
    ModuleVersion     = '1.0.1'
    Author            = 'Jeremy Dion'
    Description       = 'Devoir pour le cour CR451 module'
    PowerShellVersion = '7.5'
    FunctionsToExport = '*'
    CmdletsToExport   = @('*')
    AliasesToExport   = @()
    #Module a importé avant l'utilisation de ce module
    RequiredModules   = @(
        @{ModuleName='Microsoft.Graph.Calendar'}
    )
}