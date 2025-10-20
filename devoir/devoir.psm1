<#commentaire pour la première fonction 
 @params :
 @return :
#>
function creerDocumentDevoir {
     $word = New-Object -ComObject word.application
     $word.Visible = $True
     $doc = $word.documents.add()

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

<#commentaire pour la deuxième fonction 
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

<#commentaire pour la troisième fonction 
 @params :
 @return :
#>
function creerBulletin {
     Write-Host("Bulletin creer");
}