$Destinataire = Read-Host "Entrez l'adresse email du destinataire"

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)

# Paramètres du mail
$Mail.To = $Destinataire
$Mail.Subject = "[SUPPORT INFORMATIQUE] - Procédure informatique"
$Mail.Body = @"
Bonjour,

Vous trouverez ci-joint la procédure informatique pour transférer vos contacts Android sur IOS.

Cordialement
Le Support Informatique
"@

# Ajout de la pièce jointe
$PieceJointe = "C:\Temp\Procedure8.pdf"

if (Test-Path $PieceJointe) {
    $Mail.Attachments.Add($PieceJointe)
} else {
    Write-Host "La pièce jointe n'a pas été trouvée : $PieceJointe" -ForegroundColor Red
    exit
}

# Envoi du mail
$Mail.Send()

Write-Host "Mail envoyé avec succès à $Destinataire" -ForegroundColor Green
