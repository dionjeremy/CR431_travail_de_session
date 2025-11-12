# commande pour importer le module devoir
# import-Module Microsoft.Graph.Calendar
# import-Module studentToolKit
# appeler votre fonction 
# si vous modifier les documents du module faite la commande d'import du studentToolKit avec le paramètre -force

<#La fonction Permet d'ouvrir l'application word présente sur le poste pour généré un la base 
  d'un document de travail.La page titre,la table de matière ainsi qu'une section pour les références
  sont créée automatiquement avec les paramètres donnés par l'utilisateur.
  @params :$lang           -> langue à appliquer lors de la génération du document, fr/en sont les choix supportés
           $marge          -> Valeur à attribuer au marge, par défaut la valeur est de 36 soit 1,27cm
           $nomCours       -> Nom du cours à afficher sur la page titre , une valeur est utilisé s'il n'est pas spécifié
           $groupe         -> Numero de groupe à afficher sur la page titre, un placeholder est utilisé si aucune valeur n'est spécifiée
           $titreTravail   -> Titre du travail à afficher sur la page titre, un placeholder est utilisé si aucune valeur n'est spécifiée
           $nomEnseignant  -> Nom de l'enseignant/chargé de cours à afficher à la page titre, un placeholder est utilisé si aucune valeur n'est spécifié
           $dateRemise     -> Date de remise du document, si aucune valeur n'est donnée alors la date courrante sera utilisée
           $nomSousSection -> Liste contenant les valeur des sous-section a générer pour le document
#>
function New-Document {
	Param(
		[string] $lang = "fr",
		[int] $marge = 36,
		[string[]] $nomsEtudiants = "Nom Prénom",
		[string] $nomCours = "CRXXX - titre du cours",
          [string] $groupe = "XX",
		[string] $titreTravail = "Titre du travail",
          [string] $nomEnseignant = "Nom Prénom",
		[string] $dateRemise = (Get-Date -Format MM-dd-yyyy),
		[string[]] $nomsSousSections
	)
     Write-Host "Début de la création du document."
     <#Vérification de la présence de l'application word sur le post, dans le cas ou l'application n'est pas présente la commande 
       affiche un log d'erreur et termine l'execution de la fonction
     #>
     try {
          $word = New-Object -ComObject word.application
          $word.Visible = $True
          $doc = $word.documents.add()
     }catch{
          Write-Error("Il semble que l'executable word ne soit pas installez sur ce poste, il est donc impossible de créer un fichier de type .docx") -ErrorAction Stop
     }
     

     #vérification du code langue passé en paramètre si le code chosi ne correspond pas a fr/en , alors la langue par défaut "fr" est utilisée
	if($lang.ToLower() -ne "fr" -and $lang.ToLower() -ne "en"){
	     Write-Host("le code de langue $lang n'est pas valide, la langue par défaut sera donc utilisée");
		$lang = "fr";
	}
     #vérification de la liste des étudiants, la limite est mise a 15 (pour le formatage), dans le cas ou la limite est dépassée la liste est mise a vide 
     #et un log d'erreur est afficher à l'utilisateur lui disant qu'il devra ajouté les noms manuellement au document
     if($nomsEtudiants.Length -gt 15){
		Write-Error("la taille de la listes de noms d'étudiant est suppérieur à 15 , la liste est mise a vide et les noms devront être entrés manuellement");
          $nomsEtudiants = ""
	}


     #création de l'objet permettant l'ajout et la manipulation du document word
     $selection = $word.Selection;
	#Ajustement des marges
	Set-Marge $doc $marge;
     #Création de la page titre
	Add-PageIntroduction $selection $nomsEtudiants $nomCours $titreTravail $dateRemise $groupe $nomEnseignant $lang;
     #Création de la table des matières si le nombres de sous sections est suppérieur a 0
     if($nomsSousSections.Length -ne 0){
          $tableDesMatieres = Add-TableDesMatieres $selection
     }
     #Ajout des sous-sections
	Add-SousSections $selection $nomsSousSections;
     #Ajout de la page de références
	Add-Bibliographie $selection $lang;
     #update de la table des matières pour afficher les sous sections du document si le nombres de sous sections est suppérieur a 0
     if($nomsSousSections.Length -ne 0){
          $tableDesMatieres.Update();
     }
     Write-Host "Fin de la génération du document."
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

<#
 @params : $IDCours        -> String contenant les codes d'identification de cours Ex. CR431
           $nomCours       -> String contenant le noms du cours
           $cheminDossier  -> String contenant le chemin ou le dossier sera créer
           $noteDePassage  -> Double qui contient la note de passage par cours
 @return:  $bulletin       -> Tableau qui contient nos informations IDCours, nomsCours et noteDePassage
#>
#Fonction qui crée notre bulletin
function New-Bulletin {
     param (
          [Parameter(Mandatory = $True)][string[]]$IDCours,
          [Parameter(Mandatory = $True)][string[]]$nomsCours,
          [string]$cheminDossier,
          [Double[]]$noteDePassage = 60
     )
     <#Verifier si le chemin est défini par l'utilisateur sinon créer le dossier dans un endroit par défaut
     en fonction du OS pour la création du fichier csv#>
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
          #Si l'utilisateur a donné un chemin, on l'utilise au lieu d'utiliser celui par defaut
          New-Item -Path $cheminDossier -ItemType Directory -Name Bulletin | Out-Null
          $cheminDossierBulletin= Join-Path -Path $cheminDossier -ChildPath "Bulletin"
          $CheminCreationCSV = Join-Path -Path $cheminDossierBulletin -ChildPath "Bulletin.csv"
     }
     #S'assurer que les id de cours et les cours ont le même nombre de champs chaque
     if ($IDCours.count -ne $nomsCours.count){
          Write-Host "Il manque des valeurs dans le parametre IDCours ou nomsCours"
          break
     }
     $bulletin = @()
     
     <#Boucle pour la création de nos objets de chaque cours en fonction des entrees
     et les sauvegarder dans notre tableau bulletin#>
     for ($i = 0; $i -lt $IDCours.count; $i++) {

          $ajoutBulletin = [PSCustomObject]@{
               IDCours = $IDCours[$i]
               Cours = $nomsCours[$i]
               NoteDePassage = $noteDePassage[$i]
               MoyenneActuelle = $null
               MoyennePourPasserParEvaluation =  $null
               Evaluation = [PSCustomObject]@{}
          }    
          $bulletin += $ajoutBulletin
     }
     Write-Host("Bulletin creer");
     #Création du fichier csv du bulletin pour utilisation ultérieur
     $bulletin | Export-Csv -Path $cheminCreationCSV | Out-Null
     return $bulletin

}

<#
 @params : $cheminDossier  -> String contenant le chemin ou le dossier sera créer
 @return : $bulletin       -> Tableau qui contient nos informations importer du CSV
#>
#Fonction qui importe notre bulletin depuis un format CSV
function Import-Bulletin {
     param (
     [Parameter(Mandatory = $True)][string]$CheminCSV
     )
     <#On essaie le chemin donneé par l'utilisateur, si il n'existe pas, on sort un erreur comme quoi le fichier
     n'a pas été trouvé#>
     try {
          $bulletinCSV = Import-Csv -Path $CheminCSV
     } catch {
          Write-error "Fichier non trouver dans $CheminCSV"
          return
     }
     <#On boucle les lignes dans le fichier csv pour importer les informations dans notre objet
     ajoutbulletin et on l'ajoute a notre tableau $bulletin #>
     $bulletin = @()
     foreach ($ligne in $bulletinCSV) {
          $ajoutbulletin = [PSCustomObject]@{
               IDCours = [string]$ligne.IDCours
               Cours = [string]$ligne.Cours
               NoteDePassage = [double]$ligne.noteDePassage
               MoyenneActuelle = $null 
               MoyennePourPasserParEvaluation =  $null
               Evaluation = [PSCustomObject]@{}
          }
          # on verifie si MoyennePourPasserPar Evaluation n'est pas vide pour importer les informations
          if ($ligne.MoyennePourPasserParEvaluation -ne ""){
               $ajoutbulletin.MoyennePourPasserParEvaluation = [double]$ligne.MoyennePourPasserParEvaluation
               }
          <# On verifie si Evaluation n'est pas vide et si ce qui est ecrit commence par @{}
          pour pouvoir importer l'objet evaluation de notre csv#>
          if ($ligne.Evaluation -ne "" -and $ligne.Evaluation.StartsWith("@{")){
               $objetEvaluation = Invoke-Expression $ligne.Evaluation
               $ajoutbulletin.Evaluation = [PSCustomObject]$objetEvaluation
               }
          $bulletin += $ajoutBulletin
     }
     write-Host "Importation du Bulletin CSV terminée"
     return $bulletin
}

<#
 @params : $bulletin       -> Tableau provenant d'une fonction antérieur
           $IDCours        -> String contenant les codes d'identification de cours Ex. CR431
           $nomEvaluation  -> String contenant le noms des évaluations
           $note           -> Double qui contient les notes à modifier
           $ponderation    -> Double qui contient la pondération de chacune des évaluations
 @return:  $bulletin       -> Tableau modifié qui contient nos informations IDCours, nomsCours, noteDePassage, 
                              MoyenneActuelle, MoyennePourPasserParEvaluation et l'objet Evaluation
#>
#Fonction qui créer notre bulletin
function Set-Bulletin {
     param (
          [Parameter(Mandatory=$True, ValueFromPipeline=$True)][psobject[]]$bulletin,
          [Parameter(Mandatory=$True)][string]$IDCours,
          [Parameter(Mandatory=$True)][string[]]$nomEvaluation,
          [Parameter(Mandatory=$True)][double[]]$note,
          [Parameter(Mandatory=$True)][double[]]$ponderation
     )
     #On vérifie que le cours qu'on souhaite modifier existe dans le csv ou variable 
     $modifiercours = $bulletin | Where-Object {$_.IDCours -eq $IDCours}
     if (-not $modifiercours) {
          Write-Error "Le cours $IDCours ne se trouve pas dans le bulletin"
          break
     }
     #On s'assure que les trois listes aient le meme nombre d'elements
     if (($nomEvaluation.count -ne $note.count) -or ($nomEvaluation.Count -ne $ponderation.count)){
          Write-Host("Il manque des valeurs dans le parametre nomEvaluation, note ou ponderation ")
          break
     }
     #on fait notre boucle pour ajouter les evaluations avec leur note et ponderation
     for ($i = 0; $i -lt $nomEvaluation.count; $i++) {
          try{
               $eval = $nomEvaluation[$i]
               $notePonderation = [PSCustomObject]@{
                    Note = $note[$i]
                    Ponderation = $ponderation[$i]
               }
               add-member -InputObject $modifiercours.Evaluation -MemberType NoteProperty -Name $eval -Value $notePonderation -ErrorAction Stop   
     
          }catch{
               Write-Warning "Impossible d'ajouter l'evaluation '$eval' Il est deja existant."
          }
     }
     $totalNotePonderer = 0.0
     $totalPonderation = 0.0
     #On boucle sur nos evaluations pour chaque cours selectionner pour faire le calcul de la note pondere
    foreach ($eval in $modifiercours.Evaluation.psobject.Properties) {
        $noteEval = $eval.Value.Note
        $pondEval = $eval.Value.Ponderation
     #on vérifie que la valeur n'est pas nul et on fait le calcul de notre note pondéré
        if ($null -ne $pondEval) {
            $totalNotePonderer += ($noteEval * $pondEval)
            $totalPonderation += $pondEval
        }
    }
    <#On vérifie si le total de la pondération est plus grand que 0, si oui, 
    on fait le calcul de notre moyenne total de nos evaluations par cours#>
    if ($totalPonderation -gt 0) {
        $moyenneActuelle = $totalNotePonderer / $totalPonderation
        $modifiercours.MoyenneActuelle = [math]::Round($moyenneActuelle, 2)
    } else {
        $modifiercours.MoyenneActuelle = $null
    }
    $noteDePassage = $modifiercours.NoteDePassage
    $ponderationRestante = 100 - $totalPonderation
    <#On vérifie si la pondération est à 0, donc que toute les évaluations ont été ajouté
     et on ajuste les champs MoyennePourPasseParEvaluation en fonction#>
    if ($ponderationRestante -eq 0) {
        if ($modifiercours.MoyenneActuelle -ge $noteDePassage) {
            $modifiercours.MoyennePourPasserParEvaluation = 0 
        } else {
            $modifiercours.MoyennePourPasserParEvaluation = $null
        }
    }
    else {
        # On calcul les points manquants pour passer avec la ponderation
        $pointsRequisPourPasser = $noteDePassage * 100
        $pointsManquants = $pointsRequisPourPasser - $totalNotePonderer
        # On vérifie si l'étudiant a déjà atteint ou dépassé la note de passage
        if ($pointsManquants -le 0) {
            $modifiercours.MoyennePourPasserParEvaluation = 0
        }
        else {
            $noteRequise = $pointsManquants / $ponderationRestante
            $noteRequiseArrondie = [math]::Round($noteRequise, 2)
            $modifiercours.MoyennePourPasserParEvaluation = $noteRequiseArrondie
        }
    }
    Write-Host "Modification du bulletin terminée"
    return $bulletin
}

<#Fonction qui permet d'ajuster les marges du document
 @params : $doc   -> Objet représentant le document word présentement ouvert
           $marge -> Valeur numéric représentant la taille des marges a appliquer au document
#>
function Set-Marge {
	Param(
		$doc,
		$marge
	)
    $doc.PageSetup.LeftMargin = $marge;
    $doc.PageSetup.RightMargin = $marge;
    $doc.PageSetup.TopMargin = $marge;
    $doc.PageSetup.BottomMargin = $marge;
}


<#Fonction pour la création des sous sections 
 @params : $selection        -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $nomsSousSections -> Liste contenant les titre de sous sections a ajouter au document
#>
function Add-SousSections {
	Param(
		$selection,
		[string[]] $nomsSousSections
	)
     Write-ElementOfArray $nomsSousSections "Heading 1"
}

<#Fonction qui créer la page titre
 @params : $selection      -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $nomsEtudiants  -> Liste contenant le noms des étudiants a ajouté a lapage titre
           $nomCours       -> String contenant le noms du cours
           $dateRemise     -> La date de remise prévu pour le document (si aucune date n'est passé la date courrante sera prise)
           $groupe         -> Numero de groupe 
           $nomEnseignant  -> Nom de l'enseignant/chargé de cours
           $lang           -> lang selectionné pour la création du document
#>
function Add-PageIntroduction {
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
     Add-Line $selection $nomCours "Title" 20 1
     

     #Ajout du numero de groupe à la page titre
     $groupeLabelFr = "Groupe : ";
     $groupeLabelEn = "Group : ";
     Add-Line $selection ((Get-LabelFromLang $lang $groupeLabelFr $groupeLabelEn) + $groupe) "Strong" 16 1
     Add-EmptyLine $selection 


     #Ajout du titre du travail à la page titre
     Add-Line $selection $titreTravail "Title" 0 1
     Add-EmptyLine $selection 


     #Ajout du noms de ou des étudiants à la page titre
     $presenteParLabelFr = "Présenté par :";
     $presenteParLabelEn = "Presented by :";
     Add-Line $selection (Get-LabelFromLang $lang $presenteParLabelFr $presenteParLabelEn) "Quote" 0 1
     Write-ElementOfArray $selection $nomsEtudiants "Quote" 1

     #Ajoute un nombre de ligne vide pour la présentation de la page titre
     Add-EmptyLine $selection (18 - $nomsEtudiants.Length)

     #Ajout du noms de ou des étudiants à la page titre
     $presenteParLabelFr = "Présenté à : ";
     $presenteParLabelEn = "Presented to : ";
     Add-Line $selection ((Get-LabelFromLang $lang $presenteParLabelFr $presenteParLabelEn) + $nomEnseignant) "normal" 0 1

     #Ajout de la Date de remise (par défault c'est la date courrante)
     $dateRemiseLabelFr = "Date de remise : ";
     $dateRemiseLabelEn = "Submitted date : ";
     Add-Line $selection ((Get-LabelFromLang $lang $dateRemiseLabelFr $dateRemiseLabelEn) + $dateRemise)"normal" 0 1
}

<#Fonction qui permet d'ajouté du texte au document et d'appliquer différent attribut
 @params : $selection      -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $text           -> Texte à ajouter au document
           $fontSize       -> taille de la police 
           $TextAlignement -> Alignement du texte (3 par default ce qui justifie le texte vers la gauche)
#>
function Add-Line {
     Param(
          $selection,
          [string] $text,
          [string] $style,
          [int] $fontSize,
          [int] $textAlignment = 3
     )
     #Verification pour la taille de la police, si elle est spécifié (non nulle ou n'est pas égale a zéro) elle est ajustée
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
function Add-SousSections {
	Param(
		$selection,
		[string[]] $nomsSousSections
	)
     Write-ElementOfArray $selection $nomsSousSections "Heading 1"
}

<#Fonction qui permet d'ajouter du texte au document avec un liste d'élément en paramètre
 @params : $selection -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $elements  -> Liste des éléments textuels à ajouter au document
           $Style     -> String qui indique le style a appliquer aux éléments
           $alignement-> int qui indique de quel façon le texte devrait être aligné dans le texte (1 correspond a du texte centrer)
#>
function Write-ElementOfArray{
     Param(
          $selection,
          [string[]] $elements,
          [string] $style,
          [int] $alignment
     )

     #Vérification pour s'assurer que la liste d'element n'est pas nulle ou vide
     if($NULL -ne $elements -and $elements.Length -ne 0 ){
          #Boucle pour ajouter du texte dans un style donnée a chaque élément de la liste
		foreach ($element in $elements) {
			if($element -ne ""){
				$selection.Style=$style
   				$selection.TypeText($element);
                    #Si l'alignement est non null alors il est ajusté 
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
function Get-LabelFromLang{
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
function Add-Bibliographie {
	Param(
		$selection,
		$lang
	)
	#Constantes pour les titres de la section bibliographie en/fr
	$titreLabelFr = "Références";
	$titreLabelEn = "References";
     Add-Line $selection (Get-LabelFromLang $lang $titreLabelFr $titreLabelEn) "Heading 1"
}

<#Fonction qui ajoute un nombre de ligne vide passé en paramètre
 @params : $selection -> Objet depuis lequel il est possible d'ajouté du contenue au document word
           $nbLigne   -> nombre de ligne vide a ajouté
#>
function Add-EmptyLine {
     param (
       $selection,
       [int] $nbLigne = 1
     )
     
     #une boucle pour ajouter un ligne vide pour un nombre de lignes données
     For($i = 0; $i -lt $nbLigne; $i++){
          Add-Line $selection "" "Normal"
     }
}

<#Fonction qui initialise la table des matière 
 @params : $selection -> Objet depuis lequel il est possible d'ajouté du contenue au document word
#>
function Add-TableDesMatieres{
     param(
          $selection
     )
     #Label pour les tables des matière en englais et en français
     $TableMatiereLabelFr = "Table des matières"
     $TableMatiereLabelEn = "Table of contents"

     #ajout du titre pour la table des matière
     Add-Line $selection (Get-LabelFromLang $lang $TableMatiereLabelFr $TableMatiereLabelEn) "Normal" 18 1
     Add-EmptyLine $selection 
     
     #initialise la table des matières avant la création des sous-sections
     $range = $selection.Range;
     return $doc.TablesofContents.add($range)
}