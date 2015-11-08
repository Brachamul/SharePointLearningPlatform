jQuery(function() {
	
	// "Nous avons presque terminé"
	$('#notificationArea').hide()

	// Contenu des onglets
	$('.ms-cui-tabContainer').hide()

	// Onglets
	$('.ms-cui-tt[data-tab]').click(function(){
		var tab = $(this).attr('data-tab')
		if ($(this).hasClass('ms-cui-tt-s')) {
			$(this).removeClass('ms-cui-tt-s')
			$('.ms-cui-tabContainer').slideUp('fast')
		} else {
			var tabName = $(this).attr('data-tab')
			activateTab(tabName)
		}
	})


	$('.o365cs-nav-contextMenu').hide()

	$('#O365_MainLink_Settings').click(function(){
		$('.o365cs-nav-contextMenu').slideToggle('fast')
	})

	$('.toggle-edit-mode').click(function() {
		if ($('body').hasClass('edit-mode')) disableEditMode()
		else enableEditMode()
		return false
	})

	$('body').prepend('<div id="solucom-modal"><div class="container"><div class="subcontainer"><img src="shupert.png"/></div><div class="content"></div></div></div></div>')

	$('#DeltaPlaceHolderMain').load('content_accueil.html')

	$('#footer').html("<strong>Notre objectif</strong> : remplacer l'interview du prédecesseur de notre client par son interview à lui !")

	$.getScript('clickHacking.js')

	openModal("<p>Bienvenue sur notre <strong>plateforme d'apprentissage de SharePoint 2013</strong> !</p><p>Prenez le temps de découvrir SharePoint, mais n'oubliez pas notre objectif : <strong>remplacer l'interview du prédecesseur de notre client par son interview à lui !</strong></p>")

})

function enableEditMode() {
	$('body').addClass('edit-mode')
	$('[data-tab="text-formatter"], [data-tab="text-inserter"]').fadeIn('fast')
	$('textarea').removeAttr('readonly')
	$('#edit-or-save-button').addClass('save')
	$('.modify-or-save-text').html('Enregistrer')
	activateTab("text-formatter")
	openModal("<p>Le fait de cliquer sur le bouton 'modifier' a déclenché le <strong>mode édition</strong>.</p><p>On peut maintenant modifier le texte de l'article.</p>")
}

function disableEditMode() {
	$('body').removeClass('edit-mode')
	$('[data-tab="text-formatter"], [data-tab="text-inserter"]').fadeOut('fast')
	$('textarea').attr('readonly')
	$('#edit-or-save-button').addClass('edit')
	$('.modify-or-save-text').html('Modifier')
	activateTab("page")
	openModal("C'est enregistré !")
	trainingComplete()
}

function openModal(text) {
	$('#solucom-modal .content').html(text)
	$('#solucom-modal').fadeIn('fast')
}

function closeModal() {
	$('#solucom-modal').fadeOut('fast')
	$('.o365cs-nav-contextMenu').slideUp('fast')
}

function activateTab(tabName) {
	$('.ms-cui-tabContainer:not([data-ribbon="' + tabName + '"])').slideUp('fast')
	$('.ms-cui-tt-s').removeClass('ms-cui-tt-s')
	$('[data-tab="' + tabName + '"]').addClass('ms-cui-tt-s')
	$('[data-ribbon="' + tabName + '"]').slideDown('fast')
}

function trainingComplete() {
	openModal("<p><strong>Félicitations !</strong></p><p>Vous avez accompli la mission, et vous connaissez <strong>un peu mieux</strong> SharePoint.</p><p>Vous pouvez maintenant continuer à explorer, ou quitter cette fenêtre.</p>")
	$('#footer').html("Mission <strong>accomplie</strong> !")
}