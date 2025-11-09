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
		[string] $nomCours = "CRXXX - titre du cours",
          [string] $groupe = "Groupe : XX",
		[string] $titreTravail = "Titre du travail",
          [string] $nomEnseignant = "Nom Prénom",
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

	if($lang.ToLower() -ne "fr" -and $lang.ToLower() -ne "en"){
		Write-Host("le code de langue $lang n'est pas valide, la langue par défaut sera donc utilisée");
		$lang = "fr";
	}

     if($nomsEtudiants.Length -gt 15){
		Write-Error("la taille de la listes de noms d'étudiant est suppérieur à 15 , la liste est mise a vide et les noms devront être entrés manuellement");
          $nomsEtudiants = ""
	}
     #création de l'objet permettant l'ajout et la manipulation du document word
     $selection = $word.Selection;

	#fonction pour la gestion des marges (par defaut la valeur de la marge est 36 ou 1,26 cm )
     Write-Host("marge")
	AjusterMarge $doc $marge;
     write-host("introduction")
	CreationPageIntroduction $selection $nomsEtudiants $nomCours $titreTravail $dateRemise $groupe $nomEnseignant $lang;
     Write-Host("Création de la table des matières");
     $tableDesMatieres = addTableDesMatieres $selection
     Write-Host("Sous-sections")
	CreationSousSections $selection $nomsSousSections;
     Write-Host("bibliographie")
	CreationBibliographie $selection $lang;
     $tableDesMatieres.Update()
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
          $dateRemise, 
          $groupe,
          $nomEnseignant,
		$lang
	)
     #Ajout du titre du cours à la page titre
     printLine $selection $nomCours "Title" 20 1
     

     #Ajout du numero de groupe à la page titre
     $groupeLabelFr = "Groupe : ";
     $groupeLabelEn = "Group : ";
     printLine $selection ((stringFromLangOption $lang $groupeLabelFr $groupeLabelEn) + $groupe) "Strong" 16 1
     addEmptyLine $selection 


     #Ajout du titre du travail à la page titre
     printLine $selection $titreTravail "Title" 0 1
     addEmptyLine $selection 


     #Ajout du noms de ou des étudiants à la page titre
     $presenteParLabelFr = "Présenté par :";
     $presenteParLabelEn = "Presented by :";
     printLine $selection (stringFromLangOption $lang $presenteParLabelFr $presenteParLabelEn) "Quote" 0 1
     iterateAndPrintArrayElement $selection $nomsEtudiants "Quote" 1

     #Ajoute un nombre de ligne vide pour la présentation de la page titre
     addEmptyLine $selection (16 - $nomsEtudiants.Length)

     #Ajout du noms de ou des étudiants à la page titre
     $presenteParLabelFr = "Présenté à : ";
     $presenteParLabelEn = "Presented to : ";
     printLine $selection ((stringFromLangOption $lang $presenteParLabelFr $presenteParLabelEn) + $nomEnseignant) "normal" 0 1

     #Ajout de la Date de remise (par défault c'est la date courrante)
     $dateRemiseLabelFr = "Date de remise : ";
     $dateRemiseLabelEn = "Submitted date : ";
     printLine $selection ((stringFromLangOption $lang $dateRemiseLabelFr $dateRemiseLabelEn) + $dateRemise)"normal" 0 1
}

function printLine {
     Param(
          $selection,
          [string] $text,
          [string] $style,
          [int] $fontSize,
          [int] $textAlignment = 3
     )
     if($NULL -ne $fontSize -and $fontSize -ne 0){
          $selection.Font.Size = $fontSize;
     }

     $selection.Style = $style
	$selection.TypeText($text);
     $selection.ParagraphFormat.Alignment = $textAlignment
     $selection.TypeParagraph();

}


<#Fonction qui ajoute les sous-sections passé en paramètre si aucune n'est passée , rien n'est ajouté 
 @params : $selection         -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $nomsSousSections  -> Liste des titres des sous-sections a ajouter au document
#>
function CreationSousSections {
	Param(
		$selection,
		[string[]] $nomsSousSections
	)
     iterateAndPrintArrayElement $selection $nomsSousSections "Heading 1"
}

<#Fonction qui permet d'ajouter du texte au document avec un liste d'élément en paramètre
 @params : $selection -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $elements  -> Liste des éléments textuels à ajouter au document
           $Style     -> String qui indique le style a appliquer aux éléments
           $alignement-> int qui indique de quel façon le texte devrait être aligné dans le texte (1 correspond a du texte centrer)
#>
function iterateAndPrintArrayElement{
     Param(
          $selection,
          [string[]] $elements,
          [string] $style,
          [int] $alignment
     )

     if($NULL -ne $elements -and $elements.Length -ne 0 ){
		foreach ($element in $elements) {
			if($element -ne ""){
				$selection.Style=$style
   				$selection.TypeText($element);
                    if($NULL -ne $alignment){
                         $selection.ParagraphFormat.Alignment = $alignment
                    }
    			     $selection.TypeParagraph();
			}	
		}

	}
}

<#Fonction pour le retour d'un label selon la langue choisie (en/fr)
 @params : $lang    -> correspond a la langue choisie par l'utilisateur
           $labelFr -> valeur du label pour texte en francais
           $labelEn -> valeur du label pour texte en englais
 @return : le label contenant le texte pour la langue passé en paramètre
#>
function stringFromLangOption{
     param(
          [string] $lang,
          [string] $labelFr,
          [string] $labelEn
     )
     if($lang.ToLower() -eq "fr"){
          return $labelFr;
     }else{
          return $labelEn;
     }
}

<#Fonction qui ajoute la section Bibliographie au document 
 @params : $selection         -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $nomsSousSections  -> Liste des titres des sous-sections a ajouter au document
#>
function CreationBibliographie {
	Param(
		$selection,
		$lang
	)
	#Constantes pour les titres de la section bibliographie en/fr
	$titreLabelEng = "Bibliography";
	$titreLabelFr = "Bibliographie";
     printLine $selection (stringFromLangOption $lang $titreLabelFr $titreLabelEng) "Heading 1"
}

function addEmptyLine {
     param (
       $selection,
       [int] $nbLigne = 1
     )
     
     For($i = 0; $i -lt $nbLigne; $i++){
          printLine $selection "" "Normal"
     }
}

function addTableDesMatieres{
     param(
          $selection
     )
     $TableMatiereLabelFr = "Table des matières"
     $TableMatiereLabelEn = "Table of contents"
     printLine $selection (stringFromLangOption $lang $TableMatiereLabelFr $TableMatiereLabelEn) "Normal" 18
     addEmptyLine $selection 
     $range = $selection.Range;
     return $doc.TablesofContents.add($range)
}