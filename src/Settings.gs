var CACHE_PROP = CacheService.getScriptCache();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var SETTINGS_SHEET = "_Settings";
var CACHE_SETTINGS = false;
var SETTINGS_CACHE_TTL = 900;
var cache = JSONCacheService();
var SETTINGS = getSettings();

function getSettings() { 
  if(CACHE_SETTINGS) {
    var settings = cache.get("_settings");
  }
  
  if(settings == undefined) {
    var sheet = ss.getSheetByName(SETTINGS_SHEET);
    var values = sheet.getDataRange().getValues();
  
    var settings = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      settings[row[0]] = row[1];
    }
    
    cache.put("_settings", settings, SETTINGS_CACHE_TTL);
  }
  //Logger.log(settings);
  return settings;
}
function JSONCacheService() {
  var _cache = CacheService.getPublicCache();
  var _key_prefix = "_json#";
  
  var get = function(k) {
    var payload = _cache.get(_key_prefix+k);
    if(payload !== undefined) {
      JSON.parse(payload);
    }
    return payload
  }
  
  var put = function(k, d, t) {
    _cache.put(_key_prefix+k, JSON.stringify(d), t);
  }
  
  return {
    'get': get,
    'put': put
  }
}