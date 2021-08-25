# MsGraphMail.psm1
$FilesToImport = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue) + @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

foreach ($ImportedFile in @($FilesToImport)){
   try
   {
       . $ImportedFile.FullName
   }
   catch
   {
       Write-Error -Message "Failed to import function $($ImportedFile.FullName): $_"
   }
}
Export-ModuleMember -Function $FilesToImport.BaseName -Alias *