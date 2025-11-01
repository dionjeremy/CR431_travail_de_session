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
function creerBulletin {
     Write-Host("Bulletin creer");
}

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

}