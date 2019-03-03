// DISCLAIMER:
// Google Apps Script can be weird and all that matters is speed :-)

//////* START OF HEADER *//////
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var request = sheet.getRange("web_scraping!A1").getValue();
var diff = sheet.getRange("web_scraping!A2").getValue();
var partition = sheet.getRange("web_scraping!A4").getValue(); 
var healers = ["Restoration", "Holy", "Discipline", "Mistweaver"];
var api_key = sheet.getRange("K1:M1").getValue();

// Corresponding rows for importHTML in data-sheet
// not pretty but fast.
var classes = {
  "Death Knight": ['A', 'B', 'C', 'D', 'E'],
  "Demon Hunter": ['G', 'H', 'I', 'J', 'K'],
  "Druid": ['M', 'N', 'O', 'P', 'Q'],
  "Hunter": ['S', 'T', 'U', 'V', 'W'],
  "Mage": ['Y', 'Z', 'AA', 'AB', 'AC'],
  "Monk": ['AE', 'AF', 'AG', 'AH', 'AI'],
  "Paladin": ['AK', 'AL', 'AM', 'AN', 'AO'],
  "Priest": ['AQ', 'AR', 'AS', 'AT', 'AU'],
  "Rogue": ['AW', 'AX', 'AY', 'AZ', 'BA'],
  "Shaman": ['BC', 'BD', 'BE', 'BF', 'BG'],
  "Warlock": ['BI', 'BJ', 'BK', 'BL', 'BM'],
  "Warrior": ['BO', 'BP', 'BQ', 'BR', 'BS'],
  "All": ['BU', 'BV', 'BW', 'BX', 'BY'],
}

// should correspond to the range you want to insert data to
var startrow = 6;
var endrow = sheet.getRange("web_scraping!C4").getValue();

// change to speed up execution time,
// stops loop at 'bosses' amount of parses found.
var bosses = sheet.getRange("LFG!R2").getValue();
bosses = Number(bosses)

//////* END OF HEADER *//////

function getData_(a) {
  var names = sheet.getRange("web_scraping!" + a[0] + startrow + ":" + a[0] + endrow).getValues();
  var guilds = sheet.getRange("web_scraping!" + a[1] + startrow + ":" + a[1] + endrow).getValues();
  var realms = sheet.getRange("web_scraping!" + a[2] + startrow + ":" + a[2] + endrow).getValues();
  var itemlvls = sheet.getRange("web_scraping!" + a[3] + startrow + ":" + a[3] + endrow).getValues();
  var times = sheet.getRange("web_scraping!" + a[4] + startrow + ":" + a[4] + endrow).getValues();
  
  insertData_(names, guilds, realms, itemlvls, times)
}

// Helper function because we should not operate cross sheets in one function (speed)
function insertData_(names, guilds, realms, itemlvls, times) {
  sheet.getRange("C" + startrow + ":C" + endrow).setValues(names);
  sheet.getRange("D" + startrow + ":D" + endrow).setValues(guilds);
  sheet.getRange("E" + startrow + ":E" + endrow).setValues(realms);
  sheet.getRange("G" + startrow + ":G" + endrow).setValues(itemlvls);
  sheet.getRange("H" + startrow + ":H" + endrow).setValues(times);
  sheet.getRange("J" + startrow + ":J" + endrow).setValue("");
  sheet.getRange("L" + startrow + ":Y" + endrow).setValue("");
  sheet.getRange("AA" + startrow + ":AA" + endrow).setValue("");
}

function clearData_() {
  sheet.getRange("C" + startrow + ":C" + endrow).setValue("");
  sheet.getRange("D" + startrow + ":D" + endrow).setValue("");
  sheet.getRange("E" + startrow + ":E" + endrow).setValue("");
  sheet.getRange("G" + startrow + ":G" + endrow).setValue("");
  sheet.getRange("H" + startrow + ":H" + endrow).setValue("");
  sheet.getRange("J" + startrow + ":J" + endrow).setValue("");
  sheet.getRange("L" + startrow + ":Y" + endrow).setValue("");
  sheet.getRange("AA" + startrow + ":AA" + endrow).setValue("");
  sheet.getRange("web_scraping!C3").setValue("0");
}

function GetLogs_() {
  var names = sheet.getRange("C" + startrow + ":C" + endrow).getValues();
  var realms = sheet.getRange("E" + startrow + ":E" + endrow).getValues();
  
  for (var r in realms) {
    realms[r][0] = realms[r][0].split(' ').join('-');
    realms[r][0] = realms[r][0].split('\'').join('');
  }
  
  for (var i = startrow; i <= endrow; i++) {
    var name = names[i-startrow]
    var realm = realms[i-startrow]
    
    if (name == "" || realm == "") {
      // empty name or realm err
      sheet.getRange("AA" + i).setValue("ERR");
      continue;
    }
    
    try {
      var response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + name + "/" + realm + "/EU?partition=" + partition + "&timeframe=historical&api_key=" + api_key);
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
          response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + name + "/" + realm + "/EU?partition="+ partition + "&metric=hps&timeframe=historical&api_key=" + api_key);
          parsed = JSON.parse(response);
        }
        
        // this eats up some time
        var timestamp = sheet.getRange("A300").getValue();
        sheet.getRange("AA" + i).setValue(timestamp);
        sheet.getRange("J" + i).setValue(spec);
        var datarange = [['','','','','','','','','','','','','','']];
        
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
    case 0: clearData_(); return;
    case 1: getData_(classes["Death Knight"]); break;
    case 2: getData_(classes["Demon Hunter"]); break;
    case 3: getData_(classes["Druid"]); break;
    case 4: getData_(classes["Hunter"]); break;
    case 5: getData_(classes["Mage"]); break;
    case 6: getData_(classes["Monk"]); break;
    case 7: getData_(classes["Paladin"]); break;
    case 8: getData_(classes["Priest"]); break;
    case 9: getData_(classes["Rogue"]); break;
    case 10: getData_(classes["Shaman"]); break;
    case 11: getData_(classes["Warlock"]); break;
    case 12: getData_(classes["Warrior"]); break;
    case 13: getData_(classes["All"]); break;
    default: sheet.getRange("web_scraping!C3").setValue("0"); return;
  }
  
  GetLogs_();
  
  sheet.getRange("web_scraping!C3").setValue("0");
  lock.releaseLock();
}
