jQuery(function() {
	
	// "Nous avons presque terminé"
	$('#notificationArea').hide()

	// Contenu des onglets
	$('.ms-cui-tabContainer').hide()

	// Onglets
	$('.ms-cui-tt[data-tab]').click(function(){
		var tab = $(this).attr('data-tab')
		console.log(tab)
		if ($(this).hasClass('ms-cui-tt-s')) {
			$(this).removeClass('ms-cui-tt-s')
			$('.ms-cui-tabContainer').slideUp('fast')
		} else {
			$('.ms-cui-tabContainer:not([data-ribbon="' + tab + '"])').slideUp('fast')
			$('.ms-cui-tt-s').removeClass('ms-cui-tt-s')
			$(this).addClass('ms-cui-tt-s')
			$('[data-ribbon="' + tab + '"]').slideDown('fast')
		}
	})


	$('.o365cs-nav-contextMenu').hide()

//	$('#O365_MainLink_Settings').click(function(){
//		$('.o365cs-nav-contextMenu').slideToggle('fast')
//	})

	$('.edit-page').click(function() {
		alert("Go mode édition")
		return false
	})

	$('body').prepend('<div id="solucom-modal"><div class="container"><div class="subcontainer"><img src="shupert.png"/><div class="content"></div></div></div></div>')

	$('#DeltaPlaceHolderMain').load('content_accueil.html')

	$('#footer').html("<strong>Notre objectif</strong> : remplacer l'interview du prédecesseur de notre client par son interview à lui !")

	$.getScript('clickHacking.js')

})