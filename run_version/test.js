function myTest() {
  // Get the value for the user property 'DISPLAY_UNITS'.
  var userProperties = PropertiesService.getUserProperties();
  var test_email = userProperties.getProperty('e-mail');
  Logger.log('e-mail: %s', test_email);
}
