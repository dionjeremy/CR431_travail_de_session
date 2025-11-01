<#commentaire pour la première fonction (résponsabilité de jeremy)
 @params :
 @return :
#>
function creerDocumentDevoir {
     $ErrorActionPreference = 'Stop';
     try {
          $word = New-Object -ComObject word.application
          $word.Visible = $True
          $doc = $word.documents.add()
     }
     catch{
         Write-Error("Il semble que l'executable word ne soit pas installez sur ce poste, il est donc impossible de créer un fichier de type .docx");
     }

     #Set les marges du document 
     $margin = 36 # 1.26 cm
     $doc.PageSetup.LeftMargin = $margin;
     $doc.PageSetup.RightMargin = $margin;
     $doc.PageSetup.TopMargin = $margin;
     $doc.PageSetup.BottomMargin = $margin;
     
     #Ajoute du texte au document
     $selection = $word.Selection;
     $selection.TypeText("Hello world!")
     $selection.TypeParagraph()

     #Sauvegarde le document word au repertoire ou la commande a été appelé 
     $filename = 'C:\Demo.docx'
     $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
     $mydoc.SaveAs([ref][system.object]$filename, [ref]$saveFormat)
     $mydoc.Close()
     $MSWord.Quit()

     #affichage à l'utilisateur que le document est créé avec succès
     write-Host("document pour le devoir creer");
}

<#commentaire pour la deuxième fonction (responsabilité Abdel)
 @params :
 @return :
#>
function creerEvenementCalendrier {

     $params = @{
	     subject = "Test création event"
	     body = @{
		     contentType = "HTML"
		     content = "test"
	     }
	     start = @{
		     dateTime = "2025-01-01T00:00:00"
		     timeZone = "Pacific Standard Time"
	     }
	     end = @{
		     dateTime = "2025-01-01T01:00:00"
		     timeZone = "Pacific Standard Time"
	     }
	     location = @{
		     displayName = ""
	     }
	     attendees = @(
		     @{
			     emailAddress = @{
				     address = ""
				     name = ""
			     }
			     type = "required"
		     }
	     )
	     transactionId = ""
     }

     #Verifier les paramètres pour la création de l'évenement
     New-MgUserCalendarEvent -UserId $userId -CalendarId $calendarId -BodyParameter $params
     Write-Host("Évenement de calendrier creer");
}

<#commentaire pour la troisième fonction (responsabilité Gabriel)
 @params :
 @return :
#>
function New-Bulletin {
     param (
          [Parameter(Mandatory = $True)][string[]]$IDCours,
          [Parameter(Mandatory = $True)][string[]]$nomsCours,
          [string]$cheminDossier,
          [Double[]]$noteDePassage = 60
     )
     <#Verifier si le chemin est defini par l'utilisateur sinon creer le dossier dans un endroit par defaut
     en fonction du os pour la creation du fichier csv#>
     if (-not $cheminDossier) {
          if ($IsWindows) {
               New-Item -Path "$HOME\Documents" -ItemType Directory -Name Bulletin | Out-Null
               $cheminDossier = "$HOME\Documents\Bulletin"
          } elseif ($IsLinux -or $IsMacOS) {
               New-Item -Path "$HOME/Documents" -ItemType Directory -Name Bulletin | Out-Null
               $cheminDossier = "$HOME/Documents/Bulletin"

          } else {
               Write-Host("Systeme d'exploitation invalide")
               break
          }
     } else {
          New-Item -Path $cheminDossier -ItemType Directory -Name Bulletin | Out-Null
     }
     #S'assurer que les id de cours et les cours ont le meme nombre de champs chaque
     if ($IDCours.count -ne $nomsCours.count){
          Write-Host("Il manque des valeurs dans le parametre IDCours ou nomsCours")
          break
     }
     $bulletin = @()
     <#Boucle pour la creation de nos tableaux de chaque cours en fonction des entrees
     et les sauvegarder dans notre variable pour les retourner#>
     
     for ($i = 0; $i -lt $IDCours.count; $i++) {

          $ajoutBulletin = [PSCustomObject]@{
               IDCours = $IDCours[$i]
               Cours = $nomsCours[$i]
               NoteDePassage = $noteDePassage[$i]
               MoyenneActuelle = $null
               NotePourPasser =  $null
               Evaluation = $null
     }
          $bulletin += $ajoutBulletin
     }
     Write-Host("Bulletin creer");
     return $bulletin


}