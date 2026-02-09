
//*************************************************************************
//** Función Javascript para solventar el Fix que se producia en el menu **
//** principal al ser responsive mobile.                                 **
//*************************************************************************

$(function() {
  $('.dropdown > a[tabindex]').keydown(function(event) {
    
    if (event.keyCode == 13) {
      $(this).dropdown('toggle');
    }
  });
  
  $('.dropdown-menu > .disabled, .dropdown-header').on('click.bs.dropdown.data-api', function(event) {
    event.stopPropagation();
  });

  $('.dropdown-submenu > a').submenupicker();
});


