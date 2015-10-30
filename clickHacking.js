jQuery(function() {
	
	$('.no-clicky').click(function(){
		var response = "Mais pourquoi donc avez-vous cliqué ici ?"
		var context = $(this).attr("data-noclicky")
		if (context) response = context
		openModal(response)
		return false
	})

	$('.no-clicky *').click(function(){ return false })

	function openModal(text) {
		$('#solucom-modal .content').html(text)
		$('#solucom-modal').fadeIn('fast')
	}

	$('#solucom-modal').click( function() {
		$(this).fadeOut('fast')
	})

	$('.go-page').click(function(){
		$('#DeltaPlaceHolderMain').load('content_page.html')
	})

	$(document).keyup(function(e) {
		if (e.keyCode == 27) $('#solucom-modal').fadeOut('fast')
	});

})