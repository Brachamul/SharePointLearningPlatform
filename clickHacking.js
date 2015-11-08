jQuery(function() {
	
	$('.no-clicky:not(".disable-clicky")').each(function(){
		var response = "Mais pourquoi donc avez-vous cliqué ici ?"
		var context = $(this).attr("data-noclicky")
		if (context) response = context
		$(this).click(function(){ openModal(response); return false })
		$(this).find('*:not(".no-clicky"):not(".disable-clicky")').click(function(){ openModal(response); return false })
	})

	$('.no-clicky *').click(function(){ return false })

	$('#solucom-modal').click( function() {
		closeModal()
	})

	$('.go-page').click(function(){
		$('#DeltaPlaceHolderMain').load('content_page.html')
	})

	$(document).keyup(function(e) {
		if (e.keyCode == 27) closeModal()
	});

	$('#O365_MainLink_Settings').click(function(){
		$(this).addClass('disable-clicky')
	})

})