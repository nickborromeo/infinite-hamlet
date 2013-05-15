$(document).ready(function(){
  $('.alert').hide();
});

$('#s-btn').click(function(){
  $('.alert').show();
  $('.alert').delay(6000).fadeOut('slow');
});
 
$('#cs-btn').click(function(){
  $('.alert').show();
  $('.alert').delay(6000).fadeOut('slow');
});
