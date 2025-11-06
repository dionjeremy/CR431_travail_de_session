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
	$HEADING_TEXT_STYLE = "Heading 1";
	$TITLE_TEXT_STYLE = "Title"


    $ErrorActionPreference = 'Stop';

    try {
        $word = New-Object -ComObject word.application
        $word.Visible = $True
        $doc = $word.documents.add()
    }catch{
        Write-Error("Il semble que l'executable word ne soit pas installez sur ce poste, il est donc impossible de créer un fichier de type .docx");
    }
	#fonction pour la gestion des marges (par defaut la valeur de la marge est 36 ou 1,26 cm )
	AjusterMarge($doc);
	AjusterStyle($doc);
	CreationPageIntroduction($doc);
	CreationSousSection($doc);
     
    #Ajoute du texte au document
    $selection = $word.Selection;
	$selection.Style="Heading 1"
    $selection.TypeText("Hello world!");
    $selection.TypeParagraph();
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
<<<<<<< HEAD


=======


function AjusterMarge {
	Param($doc)
	#Set les marges du document 
    $margin = 36 # 1.26 cm
    $doc.PageSetup.LeftMargin = $margin;
    $doc.PageSetup.RightMargin = $margin;
    $doc.PageSetup.TopMargin = $margin;
    $doc.PageSetup.BottomMargin = $margin;
}

function AjusterStyle {
	Param($doc)
	
}

function CreationPageIntroduction {
	Param($doc)

}

function CreationSousSection {
	Param($doc)
>>>>>>> c0869eabc1472f399ebdeabdbbcab1efe2e45e3c

}