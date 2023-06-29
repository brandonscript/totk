function Singleton(cb, cacheKey = "function"){
  var scriptProperties = PropertiesService.getScriptProperties();
  var running = scriptProperties.getProperty(cacheKey);
  if(running.toString() === "true") {
    console.log(`Function ${cacheKey} is already running and can't be run again`)
    return;
  }
  scriptProperties.setProperty(cacheKey, "true"); // Apps Script is maniacal and won't store these as bools
  cb?.();
  scriptProperties.setProperty(cacheKey, "false");
}
