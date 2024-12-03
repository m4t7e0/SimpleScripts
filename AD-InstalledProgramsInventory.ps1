# Importa i moduli di Active Directory PowerShell
Import-Module ActiveDirectory

# Ottieni l'elenco dei computer dalla Active Directory
$computers = Get-ADComputer -Filter * -Properties Name | Select-Object -ExpandProperty Name

# Crea un file per salvare i risultati nella cartella Documenti dell'utente corrente
$outputFile = "$env:USERPROFILE\Documents\$(Get-Date -Format 'yyyy-MM-dd')-Inventory.csv"
"ComputerName,ProgramName,Version" | Out-File -FilePath $outputFile

foreach ($computer in $computers) {
    try {
        # Ottieni l'elenco dei programmi installati
        $installedPrograms = Get-WmiObject -Class Win32_Product -ComputerName $computer

        foreach ($program in $installedPrograms) {
            # Scrivi i risultati nel file CSV
            "$($computer),$($program.Name),$($program.Version)" | Out-File -FilePath $outputFile -Append
        }
    } catch {
        # Gestisci eventuali errori
        "$computer,ERROR: $($_.Exception.Message)" | Out-File -FilePath $outputFile -Append
    }
}
