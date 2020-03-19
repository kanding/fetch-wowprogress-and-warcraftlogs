// DISCLAIMER:
// Google Apps Script can be weird and all that matters is speed :-)

//////* START OF HEADER *//////
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var request = sheet.getRange("web_scraping!A1").getValue();
var diff = sheet.getRange("web_scraping!A2").getValue();
var partition = sheet.getRange("web_scraping!A4").getValue(); 
var healers = ["Restoration", "Holy", "Discipline", "Mistweaver"];
var api_key = sheet.getRange("K1:M1").getValue();

// Corresponding cols for importHTML in data-sheet
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
var clearrows = 300

// change to speed up execution time,
// stops loop at 'bosses' amount of parses found.
var bosses = Number(sheet.getRange("LFG!R2").getValue());

//////* END OF HEADER *//////

function insertSheetmsg_(range, msg) {
  sheet.getRange(range).setValue(msg);
  SpreadsheetApp.flush();
}

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
}

function clearData_() {
  sheet.getRange("C" + startrow + ":C" + clearrows).setValue("");
  sheet.getRange("D" + startrow + ":D" + clearrows).setValue("");
  sheet.getRange("E" + startrow + ":E" + clearrows).setValue("");
  sheet.getRange("G" + startrow + ":G" + clearrows).setValue("");
  sheet.getRange("H" + startrow + ":H" + clearrows).setValue("");
  sheet.getRange("J" + startrow + ":J" + clearrows).setValue("");
  sheet.getRange("L" + startrow + ":Y" + clearrows).setValue("");
  sheet.getRange("AA" + startrow + ":AA" + clearrows).setValue("");
  sheet.getRange("O3").setValue("");
}

function FetchLogs_(names, realms) {
  for (var i = startrow; i <= endrow; i++) {
    var name = names[i-startrow]
    var realm = realms[i-startrow]
    var fetch_attempts = 3
    var fetch_wait = 8000
    var fetching = "."
    
    if (name == "" || realm == "") {
      // empty name or realm err
      sheet.getRange("AA" + i).setValue("ERR");
      continue;
    }
    
    fetch : {
      while (fetch_attempts > 0) {
        try {
          var response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + name + "/" + realm + "/EU?partition=" + partition + "&timeframe=historical&api_key=" + api_key, {muteHttpExceptions: true});
          var rcode = response.getResponseCode();
          if (rcode != 200) { throw rcode }
        } catch(e) {
          Logger.log(e)
          var response = false
          // we only care about 'Too many requests.' error
          if (rcode != 429) {
            sheet.getRange("AA" + i).setValue("IMP");
            break fetch;
          }
          
          insertSheetmsg_("O3", "Too many API requests. Will resume shortly...")
          Logger.log("Fetch failed..." + fetch_attempts)
          
          Utilities.sleep(fetch_wait)
          
          fetch_attempts -= 1
          insertSheetmsg_("O3", "")
          if (fetch_attempts == 0) {
            // import error
            sheet.getRange("AA" + i).setValue("IMP");
            break;
          }
        }
        
        if (response) {
          InsertLogs_(name, realm, response, i)
          break fetch;
        }
      }
    }
  }
}


function InsertLogs_(name, realm, response, row) {
  var parsed = JSON.parse(response);
  if (parsed != "" && !parsed.hidden) {
    // this eats up some time
    var timestamp = sheet.getRange("LFG!A300").getValue();
    sheet.getRange("AA" + row).setValue(timestamp);
    
    var spec = parsed[0]["spec"];
    var class = parsed[0]["class"];
    if (healers.indexOf(spec) != -1.0) {
      response = UrlFetchApp.fetch("https://www.warcraftlogs.com:443/v1/rankings/character/" + name + "/" + realm + "/EU?partition="+ partition + "&metric=hps&timeframe=historical&api_key=" + api_key);
      parsed = JSON.parse(response);
    }
    
    var datarange = [['','','','','','','','','','','','','','']];
    var j = 0;
    var previous = ""
    var dirty = false
    
    for (var key in parsed) {
      if (j < bosses && parsed[key]["encounterName"] && parsed[key]["difficulty"] == diff) {
        if (parsed[key]["spec"] == spec) {
          datarange[0][j] = (parsed[key]["percentile"]).toFixed(1);
          previous = parsed[key]["encounterName"];
          j++;
        } else if (parsed[key]["encounterName"] == previous && parsed[key]["percentile"] > datarange[0][j-1]) {
          datarange[0][j-1] = (parsed[key]["percentile"]).toFixed(1);
          dirty = true;
        }
      }
    }
    
    if (dirty) {
      sheet.getRange("J" + row).setValue(class);
    } else {
      sheet.getRange("J" + row).setValue(spec);
    }
    sheet.getRange("L" + row + ":Y" + row).setValues(datarange);
  } else if (parsed.hidden) {
    // hidden logs
    sheet.getRange("AA" + row).setValue("HID");
  } else {
    // empty json error
    sheet.getRange("AA" + row).setValue("ERR");
  }
}

function FillLFG() {
  var lock = LockService.getScriptLock();
  lock.waitLock(1000);
  
  if (api_key == "") {
    lock.releaseLock();
    return;
  }
  
  clearData_()
  sheet.getRange("web_scraping!C3").setValue("1");
  
  switch(request) {
    case 0: 
      clearData_();
      sheet.getRange("web_scraping!C3").setValue("0");
      return;
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
  
  var names = sheet.getRange("C" + startrow + ":C" + endrow).getValues();
  var realms = sheet.getRange("E" + startrow + ":E" + endrow).getValues();
  
  for (var r in realms) {
    realms[r][0] = realms[r][0].split(' ').join('-');
    realms[r][0] = realms[r][0].split('\'').join('');
  }
  
  FetchLogs_(names, realms)
  
  sheet.getRange("web_scraping!C3").setValue("0");
  lock.releaseLock();
}
