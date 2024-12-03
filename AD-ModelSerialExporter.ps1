# Importa il modulo Active Directory
Import-Module ActiveDirectory

# Ottenere tutti i computer dal dominio
$computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name

# Creazione di un file per salvare i risultati
$outputFile = "$env:USERPROFILE\Documents\$(Get-Date -Format 'yyyy-MM-dd')-ModelSerialExported.csv" 
"ComputerName,SerialNumber,Model" | Out-File -FilePath $outputFile

foreach ($computer in $computers) {
    try {
        $bios = Get-WmiObject -Class Win32_Bios -ComputerName $computer
        $cs = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer
        $serialNumber = $bios.SerialNumber
        $model = $cs.Model
        
        # Scrittura dei dati nel file CSV
        "$computer,$serialNumber,$model" | Out-File -FilePath $outputFile -Append
    } catch {
        # Gestione degli errori nel caso il computer non sia raggiungibile
        "$computer,ERROR,ERROR" | Out-File -FilePath $outputFile -Append
    }
}
