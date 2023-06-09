function myFuntion(){
}

function doGet(){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetLineNotifyTokens = app.getSheetByName("Line Notify Tokens");  //連結「Line Notify Tokens」工作表
    var sheetLineNotifyTokenslastRow = sheetLineNotifyTokens.getLastRow();  //取得 Line Notify Tokens 工作表最後一列位置;
    var sheetLineNotifyTokensData = sheetLineNotifyTokens.getRange(2, 3, sheetLineNotifyTokenslastRow - 1, 1).getValues();
    let result = [];
    for (var i = 0; i < sheetLineNotifyTokensData.length; i++) {
        result.push(sheetLineNotifyTokensData[i][0]);
    }
    return ContentService.createTextOutput(result);
}

function doPost(e){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetLineNotifyTokens = app.getSheetByName("Line Notify Tokens");  //連結「Line Notify Tokens」工作表
    var sheetLineNotifyTokenslastRow = sheetLineNotifyTokens.getLastRow();  //取得 Line Notify Tokens 工作表最後一列位置;
    var sheetLineNotifyTokensData = sheetLineNotifyTokens.getRange(2, 1, sheetLineNotifyTokenslastRow - 1, 1).getValues();
    var parameter = e.parameter;
    var userName = parameter.userName;
    var location = parameter.location;
    var userId = parameter.userId;
    for(var i = 2; i < sheetLineNotifyTokenslastRow - 1; i++){
        if(userName == sheetLineNotifyTokens.getRange(i, 3).getValues()){
            sheetLineNotifyTokens.getRange(i, 9).setValue(location);            
            sheetLineNotifyTokens.getRange(i, 8).setValue(userId);            
        }
    }
}

function get_weather(location){
  
  var response = UrlFetchApp.fetch("https://opendata.cwb.gov.tw/api/v1/rest/datastore/F-C0032-001?Authorization={APIKEY}" + "&" + "locationName=" + location );
  var weatherData = JSON.parse(response);
  //Logger.log(response);
  return weatherData;
  
}
function weather_forecast(Weatherfile,location) {

var payload = {
        "method": "sendMessage",
        'text': '測試傳送天氣預報'
      }
      payload.text = location+ " 36hr天氣預報\n" + "\n" +
      Weatherfile.records.location[0].weatherElement[0].time[0].startTime+ "~" + Weatherfile.records.location[0].weatherElement[0].time[0].endTime +"\n"+
      "天氣狀態："+ Weatherfile.records.location[0].weatherElement[0].time[0].parameter.parameterName+ "\n" + 
      "降雨機率："+ Weatherfile.records.location[0].weatherElement[1].time[0].parameter.parameterName+ "%" + "\n" +
      "最低溫度："+ Weatherfile.records.location[0].weatherElement[2].time[0].parameter.parameterName+ "°C" + "\n" +
      "最高溫度："+ Weatherfile.records.location[0].weatherElement[4].time[0].parameter.parameterName+ "°C" + "\n" +
      "天氣舒適度："+ Weatherfile.records.location[0].weatherElement[3].time[0].parameter.parameterName+ "\n" + "\n" +
      Weatherfile.records.location[0].weatherElement[0].time[1].startTime+ "~" + Weatherfile.records.location[0].weatherElement[0].time[1].endTime +"\n"+
      "天氣狀態："+ Weatherfile.records.location[0].weatherElement[0].time[1].parameter.parameterName+ "\n" +
      "降雨機率："+ Weatherfile.records.location[0].weatherElement[1].time[1].parameter.parameterName+ "%" + "\n" +
      "最低溫度："+ Weatherfile.records.location[0].weatherElement[2].time[1].parameter.parameterName+ "°C" + "\n" +
      "最高溫度："+ Weatherfile.records.location[0].weatherElement[4].time[1].parameter.parameterName+ "°C" + "\n" +
      "天氣舒適度："+ Weatherfile.records.location[0].weatherElement[3].time[1].parameter.parameterName+ "\n" + "\n" +
      Weatherfile.records.location[0].weatherElement[0].time[2].startTime+ "~" + Weatherfile.records.location[0].weatherElement[0].time[2].endTime +"\n"+
      "天氣狀態："+ Weatherfile.records.location[0].weatherElement[0].time[2].parameter.parameterName+ "\n" +
      "降雨機率："+ Weatherfile.records.location[0].weatherElement[1].time[2].parameter.parameterName+ "%" +"\n"+
      "最低溫度："+ Weatherfile.records.location[0].weatherElement[2].time[2].parameter.parameterName+ "°C" + "\n" + 
      "最高溫度："+ Weatherfile.records.location[0].weatherElement[4].time[2].parameter.parameterName+ "°C" + "\n" +
      "天氣舒適度："+ Weatherfile.records.location[0].weatherElement[3].time[2].parameter.parameterName+ "\n" +"";    
      return payload;
}

function checkTime(){
    var DateTime =  new Date(new Date().toLocaleString("en-US", {timeZone: "Asia/Taipei"}));
    console.log(DateTime);
    var startTime = new Date(DateTime.getFullYear(), DateTime.getMonth(), DateTime.getDate(), checkTime[0][0], checkTime[0][1], 0);
    var endTime = new Date(DateTime.getFullYear(), DateTime.getMonth(), DateTime.getDate(), checkTime[1][0], checkTime[1][1], 0);
    if (DateTime < startTime || DateTime > endTime) { return; }
}

function checkLocation(notifyLoction, datasetDescription, capid){
    if(notifyLoction == '' || notifyLoction == "所有縣市") return true;
    if(datasetDescription == "大雷雨即時訊息"){
        notifyLoction = notifyLoction.toString().split(",");
        var area = capid.split(",");
        for(var i = 0; i < area.length; i++){
            if(notifyLoction.includes(area[i])) return true;
        }
        return false;
    }
    var response = "";
    var weatherData = "";
    if(capid == ""){
        var response = UrlFetchApp.fetch("https://opendata.cwb.gov.tw/api/v1/rest/datastore/W-C0033-001?Authorization={APIKEY}&locationName=" + notifyLoction);
        var weatherData = JSON.parse(response);
        for(var i = 0; i < weatherData.records.location.length; i++){
            if(weatherData.records.location[i].hazardConditions.hazards != ''){
                for(var j = 0; j < weatherData.records.location[i].hazardConditions.hazards.length; j++){
                    var phenomena = weatherData.records.location[i].hazardConditions.hazards[j].info.phenomena + "特報";
                    if(phenomena == datasetDescription) return true;
                }    
            }
        }
    }else{
        console.log(notifyLoction.toString().split(","));
        notifyLoction = notifyLoction.toString().split(",");
        response = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/dump/datastore?apikey={APIKEY}&capid=" + capid + "&format=json");
        weatherData = JSON.parse(response);
        if(weatherData.alert.info.area == null){
            for(var i = 0; i < weatherData.alert.info[0].area.length; i++){                
                for(var j = 0; j < notifyLoction.length; j++){
                    if(weatherData.alert.info[0].area[i].areaDesc.indexOf(notifyLoction[j]) != -1){
                        return true;
                    }
                }
            }
        }else{
            for(var i = 0; i < weatherData.alert.info.area.length; i++){                
                for(var j = 0; j < notifyLoction.length; j++){
                    if(weatherData.alert.info.area[i].areaDesc.indexOf(notifyLoction[j]) != -1){
                        return true;
                    }
                }
            }
        }
        
    }
    return false;
}

function updateCapid(weatherData, capCode, capid, row){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetRecords = app.getSheetByName("偵測紀錄");
    capid = capid.split(",");
    console.log(capid);
    var newCapid = "";
    for(var i = 0; i < capid.length; i++){
        for(var j = 0; j < weatherData.result.length; j++){
            if(weatherData.result[j].capid == capid[i]){
                if(newCapid != ""){
                    newCapid += ',';
                    newCapid += capid[i];
                }else{
                    newCapid += capid[i];
                }
                break;              
            }
        }
    }
    sheetRecords.getRange(row+2, 3).setValue(newCapid);
    //console.log(newCapid);
}

//CWB 天氣警特報
function cwb_Alert(){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetRecords = app.getSheetByName("偵測紀錄");
    var prevData = sheetRecords.getRange(2, 1, 6 , 2).getValues();
    var response = UrlFetchApp.fetch("https://opendata.cwb.gov.tw/api/v1/rest/datastore/W-C0033-002?Authorization={APIKEY}");
    var weatherData = JSON.parse(response);
    //console.log(weatherData.records.record);
    if(weatherData.records.record[0] != null){
        for(var i = 0; i < weatherData.records.record.length; i++){
            var datasetDescription = weatherData.records.record[i].datasetInfo.datasetDescription;
            //console.log(datasetDescription);
            var description = weatherData.records.record[i].contents.content.contentText;
            var startIndex = 0;
            var lastIndex = description.length;
            for(var j = 2; j < description.length; j++){
                if(description.charAt(j) == ' ' && description.charAt(j+1) != ' '){
                    startIndex = j+1;
                    break;
                }
            }
            if(description.indexOf("\n") == -1){
                if(description.lastIndexOf("。") != -1){
                    lastIndex = description.lastIndexOf("。") + 1;
                }
            }else{   
                if(description.indexOf("\n") == description.lastIndexOf("\n")){
                    lastIndex = description.lastIndexOf("。") + 1;
                }else{
                    lastIndex = description.lastIndexOf("\n"); 
                }                              
            }
            description = description.substring(startIndex, lastIndex);
            var validTime = weatherData.records.record[i].datasetInfo.validTime.startTime;
            validTime = new Date(validTime);
            var month = (parseInt(validTime.getMonth())+1) < 10 ? '0' + (parseInt(validTime.getMonth())+1).toString() : (parseInt(validTime.getMonth())+1).toString();
            var date = validTime.getDate() < 10 ? '0' + validTime.getDate() : validTime.getDate();
            var hour = validTime.getHours() < 10 ? '0' + validTime.getHours() : validTime.getHours();
            var minute = validTime.getMinutes() < 10 ? '0' + validTime.getMinutes() : validTime.getMinutes();
            validTime = validTime.getFullYear() + "/" + month + "/" + date + " " + hour + ":" + minute;
            //console.log(validTime);
            if(weatherData.records.record[i].hazardConditions != null){
                var phenomena = weatherData.records.record[i].hazardConditions.hazards.hazard.info.phenomena;           
                var location = "";
                for(var j = 0; j < weatherData.records.record[i].hazardConditions.hazards.hazard.info.affectedAreas.location.length; j++){
                    location += weatherData.records.record[i].hazardConditions.hazards.hazard.info.affectedAreas.location[j].locationName;
                    location += " ";
                }
                var message = "\n氣象局發佈【" + datasetDescription + "】，這些地區的朋友多留意\n\n" + "發佈時間：" + validTime + "\n發佈原因：" + description + "\n\n" + phenomena + "特報\n影響範圍：" + location;
                var image = "";
                if(datasetDescription == "大雨特報" || datasetDescription == "豪雨特報" || datasetDescription == "大豪雨特報" || datasetDescription == "超大豪雨特報"){  
                    validTime = weatherData.records.record[i].datasetInfo.validTime.startTime;
                    image = "https://www.cwb.gov.tw/Data/warning/W26_C.png?" + validTime.substring(0, 10) + validTime.substring(11,);
                }else if(datasetDescription == "陸上強風特報"){
                    validTime = weatherData.records.record[i].datasetInfo.validTime.startTime;
                    image = "https://www.cwb.gov.tw/Data/warning/W25_C.png?" + validTime.substring(0, 10) + validTime.substring(11,);
                }
                console.log(image);
                for(var k = 0; k < prevData.length; k++){
                    if(datasetDescription == prevData[k][0]){
                        if(description != prevData[k][1]){
                            var removeWarning = false;
                            sendMessege(message, image, datasetDescription, removeWarning, "");
                            sheetRecords.getRange(k+2, 2).setValue(description);
                        } 
                    }
                }
            }else{
                var message = "\n氣象局解除特報\n"+ "發佈時間：" + validTime + "\n解除原因：" + description;
                  for(var k = 0; k < prevData.length; k++){
                    if(datasetDescription == prevData[k][0]){
                        if(description != prevData[k][1]){
                            var removeWarning = true;
                            sendMessege(message, image, datasetDescription, removeWarning, "");
                            sheetRecords.getRange(k+2, 2).setValue(description);
                        } 
                    }
                }
            }
        }
    }
}

//NCDR CWB 所有警示項目
function ncdr_Alert(){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetRecords = app.getSheetByName("偵測紀錄");    
    var response = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/datastore?apikey={APIKEY}&format=json&govcode=CWB");
    var weatherData = JSON.parse(response);
    console.log(weatherData);
    if(weatherData != null && weatherData.result[0] != null){
        for(var i = 0; i < weatherData.result.length; i++){
            var prevData = sheetRecords.getRange(2, 1, 13, 3).getValues();
            var capCode = weatherData.result[i].capCode;
            var capid = weatherData.result[i].capid;
            //console.log(capCode);
            var description = weatherData.result[i].description;
            var startIndex = 0;
            var lastIndex = description.length;
            if(description.indexOf("\n") == -1){
                if(description.lastIndexOf("。") != -1){
                    lastIndex = description.lastIndexOf("。") + 1;
                }
            }else{
                if(description.indexOf("\n") != 0){
                  startIndex = 0;
                  lastIndex = description.lastIndexOf("\n"); 
                }else{
                    startIndex = description.indexOf("\n") + 1;
                    if(description.indexOf("\n") == description.lastIndexOf("\n")){
                        lastIndex = description.lastIndexOf("。") + 1;
                    }else{
                        lastIndex = description.lastIndexOf("\n"); 
                    }      
                }               
            }
            description = description.substring(startIndex, lastIndex);
            //console.log(description);
            var validTime = weatherData.result[i].effective;
            validTime = new Date(validTime);
            var month = (parseInt(validTime.getMonth())+1) < 10 ? '0' + (parseInt(validTime.getMonth())+1).toString() : (parseInt(validTime.getMonth())+1).toString();
            var date = validTime.getDate() < 10 ? '0' + validTime.getDate() : validTime.getDate();
            var hour = validTime.getHours() < 10 ? '0' + validTime.getHours() : validTime.getHours();
            var minute = validTime.getMinutes() < 10 ? '0' + validTime.getMinutes() : validTime.getMinutes();
            validTime = validTime.getFullYear() + "/" + month + "/" + date + " " + hour + ":" + minute;
            //console.log(validTime);
            if(capCode == "th"){
                var storeResponse = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/dump/datastore?apikey=={APIKEY}&capid=" + capid).getContentText();
                var weatherDataStore = XmlService.parse(storeResponse);
                var root = weatherDataStore.getRootElement();
                var xmlns = XmlService.getNamespace("urn:oasis:names:tc:emergency:cap:1.2");
                var info = root.getChildren("info", xmlns);
                var parameter = "";
                var location = "";
                var area = "";
                for(i in info){
                    var datasetDescription = info[i].getChildText("headline", xmlns);
                    parameter = info[i].getChildren("parameter", xmlns);
                    for(j in parameter){
                        if(parameter[j].getChildText("valueName", xmlns) == "townships"){
                            location = parameter[j].getChildText("value", xmlns);
                        }else if(parameter[j].getChildText("valueName", xmlns) == "counties"){
                            area = parameter[j].getChildText("value", xmlns);
                        }
                        //console.log(location); 
                    }
                }
                var datasetDescription = "大雷雨即時訊息";
                var message = "\n氣象局發佈【" + datasetDescription + "】，這些地區的朋友多留意\n\n" + "發佈時間：" + validTime + "\n發佈原因：" + description + "\n\n影響範圍:" + location;
                console.log(message);
                var image = "";
                //---------------------------------------------------------------------
                for(var k = 0; k < prevData.length; k++){
                    if(datasetDescription == prevData[k][0]){                   
                        prevCapid = prevData[k][2].split(",");
                        //console.log(prevCapid);
                        //console.log(prevCapid.includes(capid));                   
                        if(description != prevData[k][1] && prevCapid.includes(capid) == false){
                            var skipOrNot = false;
                            sendMessege(message, image, datasetDescription, skipOrNot, area);
                            sheetRecords.getRange(k+2, 2).setValue(description);
                            var newCapid = prevData[k][2];
                            if(newCapid == ''){
                                newCapid += capid;
                            }else{
                                newCapid += ',';
                                newCapid += capid;
                            }
                            //console.log(newCapid);
                            sheetRecords.getRange(k+2, 3).setValue(newCapid);
                            updateCapid(weatherData, capCode, newCapid, k);
                        }                     
                        break;
                    }
                }
                break;
            }else{
                var storeResponse = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/dump/datastore?apikey={APIKEY}&capid=" + capid + "&format=json");
                var weatherDataStore = JSON.parse(storeResponse);
                var datasetDescription = weatherDataStore.alert.info.headline;
                if(weatherDataStore.alert.info.headline == null){
                    datasetDescription = weatherDataStore.alert.info[0].headline;
                    description = "";
                    if(datasetDescription == '低溫特報'){
                        description += weatherDataStore.alert.info[0].description; 
                    }else{
                        for(var j = 0; j < weatherDataStore.alert.info.length; j++){
                          console.log(description)
                            description += weatherDataStore.alert.info[j].description; 
                        }     
                    }
                }
                //console.log(datasetDescription);
                var message = "\n氣象局發佈【" + datasetDescription + "】，這些地區的朋友多留意\n\n" + "發佈時間：" + validTime + "\n發佈原因：" + description;
                console.log(message);
                if(datasetDescription == "大雨特報" || datasetDescription == "豪雨特報" || datasetDescription == "大豪雨特報" || datasetDescription == "超大豪雨特報" || datasetDescription == "濃霧特報" || datasetDescription == "陸上強風特報"){
                    break;
                }
                var image = "";
                if(datasetDescription == "高溫資訊"){
                    image = "https://www.cwb.gov.tw/Data/warning/W29_C.png";
                }else if(datasetDescription == "海上陸上颱風警報"){
                    image = "https://www.cwb.gov.tw/Data/warning/W21_C.png";
                }    
                //----------------------------------------------------------------------------------------------------------------------
                for(var k = 0; k < prevData.length; k++){
                    if(datasetDescription == prevData[k][0]){                   
                        prevCapid = prevData[k][2].split(",");
                        //console.log(prevCapid);
                        //console.log(prevCapid.includes(capid));                   
                        if(description != prevData[k][1] && prevCapid.includes(capid) == false){
                            var skipOrNot = false;
                            if(datasetDescription != "海上颱風警報" && datasetDescription != "海上陸上颱風警報"){
                                skipOrNot = true;
                            }
                            sendMessege(message, image, datasetDescription, skipOrNot, capid);
                            sheetRecords.getRange(k+2, 2).setValue(description);
                            var newCapid = prevData[k][2];
                            if(newCapid == ''){
                                newCapid += capid;
                            }else{
                                newCapid += ',';
                                newCapid += capid;
                            }
                            //console.log(newCapid);
                            sheetRecords.getRange(k+2, 3).setValue(newCapid);
                            updateCapid(weatherData, capCode, newCapid, k);
                        }                     
                        break;
                    }
                }
            }
        }
    }
}

//淹水警戒
function ncdr_FL_Alert(){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetRecords = app.getSheetByName("偵測紀錄");
    var response = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/datastore?apikey={APIKEY}&format=json&capcode=FL&govcode=WRA");
    var weatherData = JSON.parse(response);
    console.log(weatherData);
    if(weatherData != null && weatherData.result[0] != null){
        for(var i = 0; i < weatherData.result.length; i++){
            var prevData = sheetRecords.getRange(2, 1, 13 , 3).getValues();
            var capCode = weatherData.result[i].capCode;
            var capid = weatherData.result[i].capid;
            //console.log(capCode);
            var description = weatherData.result[i].description;
            var startIndex = 0;
            var lastIndex = 2;
            if(description.indexOf("\n") == -1){
                lastIndex = description.indexOf("。") + 1;
            }else{   
                startIndex = description.indexOf("\n") + 1;
                lastIndex = description.lastIndexOf("\n");                
            }
            description = description.substring(startIndex, lastIndex);
            //console.log(description);
            var validTime = weatherData.result[i].effective;
            validTime = new Date(validTime);
            var month = (parseInt(validTime.getMonth())+1) < 10 ? '0' + (parseInt(validTime.getMonth())+1).toString() : (parseInt(validTime.getMonth())+1).toString();
            var date = validTime.getDate() < 10 ? '0' + validTime.getDate() : validTime.getDate();
            var hour = validTime.getHours() < 10 ? '0' + validTime.getHours() : validTime.getHours();
            var minute = validTime.getMinutes() < 10 ? '0' + validTime.getMinutes() : validTime.getMinutes();
            validTime = validTime.getFullYear() + "/" + month + "/" + date + " " + hour + ":" + minute;
            //console.log(validTime);
            var storeResponse = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/dump/datastore?apikey={APIKEY}&capid=" + capid + "&format=json");
            var weatherDataStore = JSON.parse(storeResponse);
            var datasetDescription = weatherDataStore.alert.info.headline;
            if(weatherDataStore.alert.info.headline == null){
                datasetDescription = weatherDataStore.alert.info[0].headline;
            }
            var message = "\n水利署發佈【" + datasetDescription + "淹水警戒】，這些地區的朋友多留意\n\n" + "發佈時間：" + validTime + "\n發佈原因：" + description;
            console.log(message);
            datasetDescription = weatherDataStore.alert.info.event;
            if(weatherDataStore.alert.info.event == null){
                datasetDescription = weatherDataStore.alert.info[0].event;
            }
            console.log(datasetDescription);
            for(var k = 0; k < prevData.length; k++){
                if(datasetDescription == prevData[k][0]){                   
                    prevCapid = prevData[k][2].split(",");
                    //console.log(prevCapid);
                    //console.log(prevCapid.includes(capid));                   
                    if(description != prevData[k][1] && prevCapid.includes(capid) == false){
                        var skipOrNot = false;
                        var image = "";
                        sendMessege(message, image, datasetDescription, skipOrNot, capid);
                        sheetRecords.getRange(k+2, 2).setValue(description);
                        var newCapid = prevData[k][2];
                        if(newCapid == ''){
                            newCapid += capid;
                        }else{
                            newCapid += ',';
                            newCapid += capid;
                        }
                        console.log(newCapid);
                        sheetRecords.getRange(k+2, 3).setValue(newCapid);
                        updateCapid(weatherData, capCode, newCapid, k);
                    }                     
                    break;
                }
            }
        }
    }
}

//空氣品質
function ncdr_AirQuality_Alert(){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetRecords = app.getSheetByName("偵測紀錄");
    var prevData = sheetRecords.getRange(2, 1, 12 , 3).getValues();
    var response = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/datastore?apikey={APIKEY}&format=json&capcode=airQuality&govcode=EPA");
    var weatherData = JSON.parse(response);
    //console.log(weatherData.result);
    if(weatherData.result[0] != null){
        for(var i = 0; i < weatherData.result.length; i++){
            var capCode = weatherData.result[i].capCode;
            var capid = weatherData.result[i].capid;
            //console.log(capCode);
            var description = weatherData.result[i].description;
            var startIndex = 0;
            var lastIndex = description.length;
            if(description.indexOf("\n") == -1){
                if(description.lastIndexOf("。") != -1){
                    lastIndex = description.lastIndexOf("。") + 1;
                }
            }else{
                startIndex = description.indexOf("\n") + 1;
                if(description.indexOf("\n") == description.lastIndexOf("\n")){
                    lastIndex = description.lastIndexOf("。") + 1;
                }else{
                    lastIndex = description.lastIndexOf("\n"); 
                }                              
            }
            description = description.substring(startIndex, lastIndex);
            console.log(description);
            var validTime = weatherData.result[i].effective;
            validTime = new Date(validTime);
            var month = (parseInt(validTime.getMonth())+1) < 10 ? '0' + (parseInt(validTime.getMonth())+1).toString() : (parseInt(validTime.getMonth())+1).toString();
            var date = validTime.getDate() < 10 ? '0' + validTime.getDate() : validTime.getDate();
            var hour = validTime.getHours() < 10 ? '0' + validTime.getHours() : validTime.getHours();
            var minute = validTime.getMinutes() < 10 ? '0' + validTime.getMinutes() : validTime.getMinutes();
            validTime = validTime.getFullYear() + "/" + month + "/" + date + " " + hour + ":" + minute;
            //console.log(validTime);
            var storeResponse = UrlFetchApp.fetch("https://alerts.ncdr.nat.gov.tw/api/dump/datastore?apikey={APIKEY}&capid=" + capid + "&format=json");
            var weatherDataStore = JSON.parse(storeResponse);
            var datasetDescription = weatherDataStore.alert.info.headline;
            if(weatherDataStore.alert.info.event == null){
                datasetDescription = weatherDataStore.alert.info[0].event;
            }
            var message = "\n環保署發佈【" + datasetDescription + "】，這些地區的朋友多留意\n\n" + "發佈時間：" + validTime + "\n發佈原因：" + description;
            console.log(message);
            for(var k = 0; k < prevData.length; k++){
                if(datasetDescription == prevData[k][0]){                   
                    prevCapid = prevData[k][2].split(",");
                    //console.log(prevCapid);
                    //console.log(prevCapid.includes(capid));                   
                    if(description != prevData[k][1] && prevCapid.includes(capid) == false){
                        var skipOrNot = true;
                        var image = "";
                        sendMessege(message, image, datasetDescription, skipOrNot, capid);
                        sheetRecords.getRange(k+2, 2).setValue(description);
                        var newCapid = prevData[k][2];
                        if(newCapid == ''){
                            newCapid += capid;
                        }else{
                            newCapid += ',';
                            newCapid += capid;
                        }
                        //console.log(newCapid);
                        sheetRecords.getRange(k+2, 3).setValue(newCapid);
                        updateCapid(weatherData, capCode, newCapid, k);
                    }                     
                    break;
                }
            }
        }
    }
}

function sendMessege(message, image, datasetDescription, skipOrNot, capid){
    var app = SpreadsheetApp.openById("1yts8EZP0qNOM_9nFEB2mHAx09vmBN6HY6A5X4uNv4Ps");
    var sheetLineNotifyTokens = app.getSheetByName("Line Notify Tokens"); //連結「Line Notify Tokens」工作表
    var sheetLineNotifyTokenslastRow = sheetLineNotifyTokens.getLastRow();  //取得 Line Notify Tokens 工作表最後一列位置;
    var sheetLineNotifyTokensData = sheetLineNotifyTokens.getRange(2, 1, sheetLineNotifyTokenslastRow - 1, 1).getValues();
    var request = [];
    //console.log(checkLocation(sheetLineNotifyTokens.getRange(3, 9).getValues()[0][0].split(","), "豪雨特報", ""));
    for (var i = 0; i < sheetLineNotifyTokensData.length; i++) {
        if(skipOrNot || checkLocation(sheetLineNotifyTokens.getRange(i+2, 9).getValues()[0][0].split(","), datasetDescription, capid)){
            var options =
            {
              "url" : "https://notify-api.line.me/api/notify",
              "muteHttpExceptions" : true,
              "method"  : "post",
              "payload" : {"message" : message,
                           "imageThumbnail":image,
                           "imageFullsize":image},
              "headers" : {"Authorization" : "Bearer " + sheetLineNotifyTokensData[i][0]}
            };
            request.push(options);
        }
    }
    UrlFetchApp.fetchAll(request);
}
