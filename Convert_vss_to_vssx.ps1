$visio = New-Object -com Visio.Application

 Get-ChildItem -Path 'C:\...\Taktische Zeichen\*.vss' -Recurse | 
 ForEach-Object {
    "Working on $($PSItem.Name)"

    $doc          = $visio.Documents.Open($PSItem.FullName)
    $vssxFileName = [io.path]::ChangeExtension($PSItem,'.vssx')

    $doc.SaveAs($VSSXFileName)
    $doc.close()
 }