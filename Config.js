// ─── CONFIG SHEET ────────────────────────────────────────────────────────────
function getConfig(key) {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('cfg_' + key);
    if (cached !== null) return cached;
    var ss    = getSS_();
    var sheet = ss.getSheetByName('Config');
    if (!sheet || sheet.getLastRow() < 1) { cache.put('cfg_' + key, '', 120); return ''; }
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0]||'').toString().trim() === key) {
        var val = (data[i][1]||'').toString().trim();
        cache.put('cfg_' + key, val, 120);
        return val;
      }
    }
    cache.put('cfg_' + key, '', 120);
    return '';
  } catch(e) { return ''; }
}

function setConfig(key, value) {
  try {
    try { CacheService.getScriptCache().remove('cfg_' + key); } catch(_) {}
    var ss    = getSS_();
    var sheet = getOrCreateSheet_(ss, 'Config');
    var data  = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0]||'').toString().trim() === key) { sheet.getRange(i+1, 2).setValue(value); return; }
    }
    sheet.appendRow([key, value]);
  } catch(e) { Logger.log('setConfig: ' + e); }
}


function isAdminEmail(email) {
  var admins = CONFIG.ADMIN_EMAILS || '';
  return admins.toLowerCase().indexOf((email||'').toLowerCase().trim()) > -1;
}
