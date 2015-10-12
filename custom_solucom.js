jQuery(function() {
	
	// "Nous avons presque terminé"
	$('#notificationArea').hide()

	// Contenu des onglets
	$('[id*="Ribbon"][role="tabpanel"]').hide()

	$('.ms-cui-tabContainer').hide()

	// Onglets
	$('[id*="Ribbon"].ms-cui-tt').click(function(){
		if ($(this).hasClass('ms-cui-tt-s')) {
			$(this).removeClass('ms-cui-tt-s')
			$('.ms-cui-tabContainer').slideUp('fast')
		} else {
			$('.ms-cui-tt-s').removeClass('ms-cui-tt-s')
			$(this).addClass('ms-cui-tt-s')
			$('[role="tabpanel"]').hide()
			$('.ms-cui-tabContainer').hide()
			id = $(this).attr("id").replace("-title", "")
			$('[id*="' + id + '"][role="tabpanel"]').show()
			$('.ms-cui-tabContainer').slideDown('fast')
		}
	})



});