var sheet = SpreadsheetApp.getActiveSpreadsheet();
var request = sheet.getRange("web_scraping!A1").getValue();
var diff = sheet.getRange("web_scraping!A2").getValue();
var partition = sheet.getRange("web_scraping!A4").getValue(); 
var healers = ["Restoration", "Holy", "Discipline", "Mistweaver"];
var api_key = sheet.getRange("K1:M1").getValue();

// change to speed up execution time
// stops after finding 'bosses' amount of percentiles
// instead of checking every key in the JSON from WCL
var bosses = sheet.getRange("LFG!R2").getValue();
bosses = Number(bosses)

function insertData_(c1, c2, c3, c4, c5) {
  var names = sheet.getRange("web_scraping!" + c1 + "6:" + c1 + "74").getValues();
  var guilds = sheet.getRange("web_scraping!" + c2 + "6:" + c2 + "74").getValues();
  var realms = sheet.getRange("web_scraping!" + c3 + "6:" + c3 + "74").getValues();
  var itemlvls = sheet.getRange("web_scraping!" + c4 + "6:" + c4 + "74").getValues();
  var times = sheet.getRange("web_scraping!" + c5 + "6:" + c5 + "74").getValues();
  sheet.getRange("Names").setValues(names);
  sheet.getRange("Guilds").setValues(guilds);
  sheet.getRange("Realms").setValues(realms);
  sheet.getRange("Item_levels").setValues(itemlvls);
  sheet.getRange("Times").setValues(times);
  sheet.getRange("Specs").setValue("");
  sheet.getRange("Logs").setValue("");
  sheet.getRange("Updates").setValue("");
}

function FillLFG() {
  var lock = LockService.getScriptLock();
  lock.waitLock(1000);
  
  if (api_key == "") {
    lock.releaseLock();
    return;
  }
  
  var running = sheet.getRange("web_scraping!C3").getValue();
  if (running == 0) {
    sheet.getRange("web_scraping!C3").setValue("1");
  }
  
  switch(request) {
    case 0:
      // None
      sheet.getRange("Names").setValue("");
      sheet.getRange("Guilds").setValue("");
      sheet.getRange("Realms").setValue("");
      sheet.getRange("Item_levels").setValue("");
      sheet.getRange("Times").setValue("");
      sheet.getRange("Specs").setValue("");
      sheet.getRange("Logs").setValue("");
      sheet.getRange("Updates").setValue("");
      sheet.getRange("web_scraping!C3").setValue("0");
      return;
    case 1:
      // Death Knight
      insertData_('A', 'B', 'C', 'D', 'E');
      break;
    case 2:
      // Demon Hunter
      insertData_('G', 'H', 'I', 'J', 'K');
      break;
    case 3:
      // Druid
      insertData_('M', 'N', 'O', 'P', 'Q');
      break;
    case 4:
      // Hunter
      insertData_('S', 'T', 'U', 'V', 'W');
      break;
    case 5:
      // Mage
      insertData_('Y', 'Z', 'AA', 'AB', 'AC');
      break;
    case 6:
      // Monk
      insertData_('AE', 'AF', 'AG', 'AH', 'AI');
      break;
    case 7:
      // Paladin
      insertData_('AK', 'AL', 'AM', 'AN', 'AO');
      break;
    case 8:
      // Priest
      insertData_('AQ', 'AR', 'AS', 'AT', 'AU');
      break;
    case 9:
      // Rogue
      insertData_('AW', 'AX', 'AY', 'AZ', 'BA');
      break;
    case 10:
      // Shaman
      insertData_('BC', 'BD', 'BE', 'BF', 'BG');
      break;
    case 11:
      // Warlock
      insertData_('BI', 'BJ', 'BK', 'BL', 'BM');
      break;
    case 12:
      // Warrior
      insertData_('BO', 'BP', 'BQ', 'BR', 'BS');
      break;
    case 13:
      // Talent scout high ilvl
      // have to change insertdata/endrow to work
      break;
    default:
      sheet.getRange("web_scraping!C3").setValue("0");
      return;
  }
  GetLogs_();
  sheet.getRange("web_scraping!C3").setValue("0");
  lock.releaseLock();
}

function GetLogs_() {
  var startrow = 6;
  var endrow = 74;
  var names = sheet.getRange("C" + startrow + ":C" + endrow).getValues();
  var realms = sheet.getRange("E" + startrow + ":E" + endrow).getValues();
  sheet.getRange("Z6:Z74").setValue("");
  sheet.getRange("Specs").setValue("");
  sheet.getRange("Logs").setValue("");
  sheet.getRange("Updates").setValue("");
  
  for (var r in realms) {
    realms[r][0] = realms[r][0].split(' ').join('-');
    realms[r][0] = realms[r][0].split('\'').join('');
  }
  
  for (var i = startrow; i <= endrow; i++) {
    if (names[i-6] == "" || realms[i-6] == "") {
      // empty name or realm err
      sheet.getRange("AA" + i).setValue("ERR");
      continue;
    }
    
    try {
      var response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + names[i-6] + "/" + realms[i-6] + "/EU?partition=" + partition + "&timeframe=historical&api_key=" + api_key);
    } catch(e) { 
      var response = false
      // import error
      sheet.getRange("AA" + i).setValue("IMP");
      continue;
    }
    
    if (response) {
      var parsed = JSON.parse(response);
      if (parsed != "" && !parsed.hidden) {
        var spec = parsed[0]["spec"];
        if (healers.indexOf(spec) != -1.0) {
          response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + names[i-6] + "/" + realms[i-6] + "/EU?partition="+ partition + "&metric=hps&timeframe=historical&api_key=" + api_key);
          parsed = JSON.parse(response);
        }
        
        var timestamp = sheet.getRange("web_scraping!C1").getValue();
        sheet.getRange("AA" + i).setValue(timestamp);
        sheet.getRange("J" + i).setValue(spec);
        var datarange = [['-','-','-','-','-','-','-','-','-','-','-','-','-','-']];
        
        // keep sequential order in datarange despite holes in JSON
        var j = 0;
        for (var key in parsed) {
          if (j < bosses && parsed[key]["encounterName"] && parsed[key]["difficulty"] == diff && parsed[key]["spec"] == spec) {
            datarange[0][j] = parsed[key]["percentile"];
            j++;
          }
        }
        sheet.getRange("L" + i + ":Y" + i).setValues(datarange);
      } else if (parsed.hidden) {
        // hidden logs
        sheet.getRange("AA" + i).setValue("HID");
      } else {
        // empty json error
        sheet.getRange("AA" + i).setValue("ERR");
      }
    }
  }
}
