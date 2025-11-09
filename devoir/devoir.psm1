#commande pour importer le module devoir
# import-Module Microsoft.Graph.Calendar
# import-Module devoir
# appeler votre fonction 
# si vous modifier les documents du module faite la commande d'import du devoir avec le paramètre -force

<#commentaire pour la première fonction (résponsabilité de jeremy)
 @params :
 @return :
#>
function creerDocumentDevoir {
	Param(
		[string] $lang = "fr",
		[int] $marge = 36,
		[string[]] $nomsEtudiants = "Nom Prénom",
		[string] $nomCours = "CRXXX - titre cours",
		[string] $titreTravail = "titre travail",
		[string] $dateRemise = (Get-Date -Format MM-dd-yyyy),
		[string[]] $nomsSousSections
	)
	write-host($nomsSousSections)

    try {
        $word = New-Object -ComObject word.application
        $word.Visible = $True
        $doc = $word.documents.add()
    }catch{
        Write-Error("Il semble que l'executable word ne soit pas installez sur ce poste, il est donc impossible de créer un fichier de type .docx");
    }

	if($lang.ToLower -ne "fr" -and $lang.ToLower -ne "en"){
		Write-Host("le code de langue $lang n'est pas valide, la langue par défaut sera donc utilisée");
		$lang = "fr";
	}
	$selection = $word.Selection;

	#fonction pour la gestion des marges (par defaut la valeur de la marge est 36 ou 1,26 cm )
	AjusterMarge $doc $marge;
	CreationPageIntroduction $selection $nomsEtudiants $nomCours $titreTravail $lang;
	CreationSousSections $selection $nomsSousSections;
	CreationBibliographie $selection $lang;
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
#a completer
function Import-Bulletin {
     param (
     [string]$CheminCSV
     )

     try {
          $bulletinCSV = Import-Csv -Path $CheminCSV
     } catch {
          Write-error "Fichier non trouver dans $CheminCSV"
          return
     }
     
     $bulletin = @()
     foreach ($ligne in $bulletinCSV) {
          $ajoutbulletin = [PSCustomObject]@{
               IDCours = [string]$ligne.IDCours
               Cours = [string]$ligne.Cours
               NoteDePassage = [double]$ligne.noteDePassage
               MoyenneActuelle = $null 
               NotePourPasser =  $null
               Evaluation = $null
          }
          
          if ($ligne.MoyenneActuelle -ne ""){
               $ajoutbulletin.MoyenneActuelle = [double]$ligne.MoyenneActuelle
               }
          if ($ligne.NotePourPasser -ne ""){
               $ajoutbulletin.NotePourPasser = [double]$ligne.NotePourPasser
               }
          if ($ligne.Evaluation -ne ""){
               $ajoutbulletin.Evaluation = [string]$ligne.Evaluation
               }
          $bulletin += $ajoutBulletin
     }
     write-Host "Importation du Bulletin CSV terminee"
     return $bulletin
}

function Set-Bulletin {
     param (

     )
}

function Get-AnalyseBulletin {
     param(

     )
}

function AjusterMarge {
	Param(
		$doc,
		$marge
	)
    $doc.PageSetup.LeftMargin = $marge;
    $doc.PageSetup.RightMargin = $marge;
    $doc.PageSetup.TopMargin = $marge;
    $doc.PageSetup.BottomMargin = $marge;
}

function CreationPageIntroduction {
	Param(
		$selection,
		$nomsEtudiants,
		$nomCours, 
		$titreTravail, 
		$lang
	)
	$selection.Font.Size = 20;
	$selection.TypeText($nomCours);
    $selection.TypeParagraph();


	$selection.Style="Title"
   	$selection.TypeText($titreTravail);
    $selection.TypeParagraph();

}

function CreationSousSections {
	Param(
		$selection,
		[string[]] $nomsSousSections
	)

	if($NULL -ne $nomsSousSections -and $nomsSousSections.Length -ne 0 ){
		foreach ($nomSousSection in $nomsSousSections) {
			if($nomSousSection -ne ""){
				$selection.Style="Heading 1"
   				$selection.TypeText($nomSousSection);
    			$selection.TypeParagraph();
			}	
		}

	}

}

function CreationBibliographie {
	Param(
		$selection,
		$lang
	)
	#Constantes pour les titres de la section bibliographie en/fr
	$titreEng = "Bibliography";
	$titreFr = "Bibliographie"

	$selection.Style="Heading 1"
	if($lang.ToLower -eq "fr" ){
		$selection.TypeText($titreFr);
	} else {
		$selection.TypeText($titreEng);
	}
    $selection.TypeParagraph();
}