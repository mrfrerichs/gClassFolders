function t(input, cachedLang) {
  var sourceLanguage = "en";
  if (cachedLang) {
    var targetLanguage = cachedLang;
  } else {
    var targetLanguage = CacheService.getPublicCache().get('lang');
  }
  if (!targetLanguage) {
    targetLanguage = ScriptProperties.getProperty('lang');
    CacheService.getPublicCache().put('lang', targetLanguage, 5);
  }
  if(!targetLanguage) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
      if (!ss) {
        ss = SpreadsheetApp.openById(this.SSKEY);
      }
      var locale = ss.getSpreadsheetLocale();
      targetLanguage = locale.substring(0,2); 
    } catch(err) {
      targetLanguage = 'en';
    }
  }
  var output = input;
  if (sourceLanguage!=targetLanguage) {
    output = LanguageApp.translate(input, sourceLanguage, targetLanguage)
  }
  return output;
}
