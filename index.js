const MY_URL = 'YOUR_CALLBACK_URL';
const MY_KEY = 'YOUR_API_KEY';
//
function onOpen(){ 
  SpreadsheetApp.getUi()
      .createMenu('Vizion Tools')
      .addItem('Subscribe to Containers','subscribeAll')
      .addItem('Unsubscribe Selection', 'unsubscribeSelection')
      .addItem('Get Last Statuses','statuses')
      .addToUi();
}

function doPost(e){ 
  cache_JSON_Post(e); 
  return ContentService.createTextOutput('ok');
}

function doGet(e){  
  var params = JSON.stringify(e);
  return HtmlService.createHtmlOutput(params);
}

const expected_location_objects = ["origin_port", "destination_port", "inland_origin", "inland_destination"];

const locationObj = {
  "name":null,
  "city":null,
  "state":null,
  "country":null,
  "unlocode":null,
  "facility":null,
  "geolocation":{
    "latitude":null,
    "longitude":null
  }
}

function getColumnFromName(columnName, sheet){
  var lc = sheet.getLastColumn();
  var headers = sheet.getRange(1,1,1,lc).getValues();
  var col = headers[0].indexOf(columnName) + 1;
  return col;
}

function formatJSON(object){ // expands null objects to prevent 'undefined' errors when flattening JSON.
  object = JSON.parse(object);
  var payload = object.payload;
  for (a in expected_location_objects){
    if (payload[expected_location_objects[a]] == null){
      payload[expected_location_objects[a]] = locationObj;
    }
    else if (payload[expected_location_objects[a]].geolocation == null){
      payload[expected_location_objects[a]].geolocation = locationObj.geolocation;
    }
  }
  for (b in payload.milestones){
    var event = payload.milestones[b];
    if (event.location == null){
      event.location = locationObj;
    }
    else if (event.location.geolocation == null){
      event.location.geolocation = locationObj.geolocation;
    }
  }
  return object;
}

function flattenJSON(object){ // creates a 2D array from JSON object
  var post = object;
  var payload = post.payload;
  const startArray = [post.reference_id, post.parent_reference_id, post.id, post.created_at, payload.container_id, payload.container_iso, payload.bill_of_lading, payload.carrier_scac];
  const endArray = [payload.origin_port.name, payload.origin_port.city, payload.origin_port.state, payload.origin_port.country, payload.origin_port.unlocode, payload.origin_port.geolocation. latitude, payload.origin_port.geolocation.longitude, payload.destination_port.name, payload.destination_port.city, payload.destination_port.state, payload.destination_port.country, payload.destination_port.unlocode, payload.destination_port.geolocation.latitude, payload.destination_port.geolocation.longitude, payload.inland_origin.name, payload.inland_origin.city, payload.inland_origin.state, payload.inland_origin.country, payload.inland_origin.unlocode, payload.inland_origin.geolocation.latitude, payload.inland_origin.geolocation.longitude, payload.inland_destination.name, payload.inland_destination.city, payload.inland_destination.state, payload.inland_destination.country, payload.inland_destination.unlocode, payload.inland_destination.geolocation.latitude, payload.inland_destination.geolocation.longitude];
  var outputArray = [];
  for (var a=0; a<payload.milestones.length; a++){
    var currentEvent = payload.milestones[a];
    var middleArray = [currentEvent.description, currentEvent.raw_description, currentEvent.timestamp, currentEvent.planned, currentEvent.source, currentEvent.mode, currentEvent.vessel, currentEvent.vessel_mmsi, currentEvent.vessel_imo, currentEvent.voyage, currentEvent.location.unlocode, currentEvent.location.name, currentEvent.location.city, currentEvent.location.state, currentEvent.location.country, currentEvent.location.facility, currentEvent.location.geolocation.latitude, currentEvent.location.geolocation.longitude];
    outputArray.push(startArray.concat(middleArray,endArray));
  }
  return outputArray;
}

function writeFlattenedJSON(jsonArray){ // writes a 2D array to the last row of the "Flattened JSON" page
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flattened JSON');
  var numRows = jsonArray.length;
  var numCols = jsonArray[0].length;
  var lastRow = sheet.getLastRow()+1;
  sheet.getRange(lastRow, 1, numRows, numCols).setValues(jsonArray);
}

function unsubscribe_from_reference(refId){ // unsubscribes from active reference and returns status message
  var options = {
    'method' : 'delete',
    'headers' : {'X-API-KEY' : MY_KEY}
  }
  var response = UrlFetchApp.fetch('https://prod.vizionapi.com/references/'+refId,options);
  var data = JSON.parse(response.getContentText());
  return data.message;
}

function subscribe_to_container(id, code, bol){ // subscribes to container or MBL and returns reference ID

  var inputArray = [id, code, bol];
  for (var x=0; x<inputArray.length; x++){
      if (inputArray[x] == null){
          inputArray[x] = "null";
      }
  }
  var input = {
    'container_id': id,
    'carrier_code': code !== '' ? code : null,
    'bill_of_lading' : bol,
    'callback_url' : MY_URL,
  };
  var options = {
    'method' : 'post',
    'headers' : {
      'X-API-Key' : MY_KEY,
      'Content-Type' : 'application/json'
      },
    'payload' : JSON.stringify(input)
  };
  var response = UrlFetchApp.fetch('https://prod.vizionapi.com/references', options);
  var data = JSON.parse(response.getContentText());
  return data.reference.id;
}

function get_all_active_reference_info(){ // returns an array of all active reference ID's (column 1), their last update attempted at (column 2), and last update status (column 3)
  var options = {
    'method' : 'get',
    'headers' : {'X-API-KEY' : MY_KEY}
  }
  var response = UrlFetchApp.fetch('https://prod.vizionapi.com/references?limit=2000', options);
  response = JSON.parse(response.getContentText());
  var activeRefs = [];
  for (var a=0; a<response.length; a++){
    var currentRef = response[a];
    activeRefs.push([currentRef.id, currentRef.last_update_attempted_at, currentRef.last_update_status])
  }
  return activeRefs;
}

function subscribeAll(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var lastRow = sheet.getLastRow();
  const containerColumn = getColumnFromName("Container ID",sheet);
  const carrierCodeColumn = getColumnFromName("Vizion Carrier Code",sheet);
  const mblColumn = getColumnFromName("MASTER BL#",sheet);
  const reference_id_Column = getColumnFromName("Reference ID",sheet);
  const dataReceivedColumn = getColumnFromName("Data Received?",sheet);
  const unsubscribeColumn = getColumnFromName("Unsubscribe?",sheet);
  for (var a=2; a<=lastRow; a++){
    if (!sheet.getRange(a,reference_id_Column).getValue()){
      var container = sheet.getRange(a,containerColumn).getValue();
      var carrier = sheet.getRange(a, carrierCodeColumn).getValue();
      var mbl = sheet.getRange(a, mblColumn).getValue();
      if (!container && !mbl){
        var zero = 0;
      }
      else {
        sheet.getRange(a, reference_id_Column).setValue(subscribe_to_container(container, carrier, mbl));
        sheet.getRange(a, dataReceivedColumn).setFormulaR1C1("=OR(COUNTIF('Flattened JSON'!R1C1:R50000C1,R[0]C[-1])>0,COUNTIF('Flattened JSON'!R1C2:R50000C2,R[0]C[-1])>0)");
        sheet.getRange(a, unsubscribeColumn).insertCheckboxes();
      }
    }
  }
}

function unsubscribeSelection(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var lastRow = sheet.getLastRow();
  const unsubscribeColumn = getColumnFromName("Unsubscribe?",sheet);
  const reference_id_Column = getColumnFromName("Reference ID",sheet);
  for (var a=2; a<=lastRow; a++){
    if (sheet.getRange(a,unsubscribeColumn).getValue()){
      var refID = sheet.getRange(a, reference_id_Column).getValue();
      unsubscribe_from_reference(refID);
      sheet.deleteRow(a);
      a=a-1;
    }
  }
}

function cache_JSON_Post(post){
  post = JSON.parse(post.postData.contents);
  var reference_id = post.reference_id;
  var cache = CacheService.getDocumentCache();
  cache.put(reference_id, JSON.stringify(post), 21600);
  if (post.parent_reference_id){
    handle_child_reference(post.payload.container_id, post.payload.bill_of_lading, reference_id)
  }
}

function writePayloads_from_cache(){
  var cache = CacheService.getDocumentCache();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var reference_id_Column = getColumnFromName("Reference ID", sheet);
  var eta_column = getColumnFromName("Event Name:", sheet);
  var updated_at_Column = getColumnFromName("Last Update Time:", sheet);
  var last_status_Column = getColumnFromName("Last Update Status:", sheet);
  var latest_update_id_Column = getColumnFromName("Latest Update ID:", sheet);
  var lastRow = sheet.getLastRow();
  var listed_references = sheet.getRange(1,reference_id_Column,lastRow,1).getValues();
  for (var a=0; a<listed_references.length; a++){
    var object = cache.get(listed_references[a][0]);
    if (object){
      var key_elems = JSON.parse(object);
      var time = key_elems.created_at;
      var status = key_elems.status;
      var update_id = key_elems.id;
      object = formatJSON(object);
      var eta = checkETAorATA(object);
      writeETAs(object)
      var arr = flattenJSON(object);
      writeFlattenedJSON(arr);
      cache.remove(listed_references[a][0]);
      sheet.getRange(a+1, updated_at_Column).setValue(time);
      sheet.getRange(a+1, last_status_Column).setValue(status);
      sheet.getRange(a+1, latest_update_id_Column).setValue(update_id);
      if (eta){
        sheet.getRange(a + 1, eta_column, 1, 2).setValues([eta.slice(0, 2)]);
      }
    }
  }
}

function handle_child_reference(container_id, bol, reference_id){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var reference_id_Column = getColumnFromName("Reference ID", sheet);
  var lastRow = sheet.getLastRow();
  var listed_references = [];
  for (var row=1; row<=lastRow; row++){
    listed_references.push(sheet.getRange(row,reference_id_Column).getValue());
  }
  var index = listed_references.indexOf(reference_id);
  if (index == -1){
    var mblColumn = getColumnFromName("MASTER BL#",sheet);
    var containerColumn = getColumnFromName("Container ID",sheet);
    var dataReceivedColumn = getColumnFromName("Data Received?",sheet);
    var unsubscribeColumn = getColumnFromName("Unsubscribe?",sheet);
    sheet.getRange(lastRow+1, reference_id_Column).setValue(reference_id);
    sheet.getRange(lastRow+1, mblColumn).setValue(bol);
    sheet.getRange(lastRow+1, containerColumn).setValue(container_id);
    sheet.getRange(lastRow+1, dataReceivedColumn).setFormulaR1C1("=OR(COUNTIF('Flattened JSON'!R1C1:R50000C1,R[0]C[-1])>0,COUNTIF('Flattened JSON'!R1C2:R50000C2,R[0]C[-1])>0)");
    sheet.getRange(lastRow+1, unsubscribeColumn).insertCheckboxes();
  }
}

function checkETAorATA(obj) {
  
  var milestones = obj.payload.milestones;

  for (var a = 0; a < milestones.length; a++) {
    var event = milestones[a];
    
    // Check if event description is not null before calling toLowerCase()
    if (event.description && typeof event.description === 'string') {
      var event_description = event.description.toLowerCase();
      
      if (event_description.indexOf("destination") > -1) {
        return [event.description, event.timestamp, event.planned];
      }
    }
  }
}

function statuses(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var last_status_column = getColumnFromName("Last Update Status:",ss);
  var reference_id_column = getColumnFromName("Reference ID", ss);
  var last_update_time_column = getColumnFromName("Last Update Time:",ss);
  var lr = ss.getLastRow();
  var activeRefs = get_all_active_reference_info();
  var refIdsOnly = [];
  for (var z=0; z<activeRefs.length; z++){
    refIdsOnly.push(activeRefs[z][0]);
  }
  var listedRefs = [];
  for (a=2; a<=lr; a++){
    listedRefs.push([ss.getRange(a,reference_id_column).getValue(), a])
  }
  for (var b=0; b<refIdsOnly.length; b++){
    var checkRef = refIdsOnly[b];
    for (var c=0; c<listedRefs.length; c++){
      if (checkRef == listedRefs[c][0]){
        ss.getRange(listedRefs[c][1],last_update_time_column).setValue(activeRefs[b][1]);
        ss.getRange(listedRefs[c][1],last_status_column).setValue(activeRefs[b][2]);
        activeRefs.splice(b,1);
        refIdsOnly.splice(b,1);
        listedRefs.splice(c,1);
        b -= 1;
        c -= 1;
      }
    }
  }
}

function archive_by_subscribed(){
  var activeRefs = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var as = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive');
  var reference_id_Column = getColumnFromName("Reference ID",ss)
  var lr = ss.getLastRow();
  for (var a=1; a<=lr; a++){
    var ref = ss.getRange(a, reference_id_Column).getValue();
    if (ref){
    activeRefs.push(ref);
    }
  }
  var fj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flattened JSON');
  lr = fj.getLastRow();
  var lc = fj.getLastColumn();
  for (var a=2;a<=lr;a++){
    if (!activeRefs.includes(fj.getRange(a,1).getValue()) && !activeRefs.includes(fj.getRange(a,2).getValue())){
        fj.getRange(a,1,1,lc).moveTo(as.getRange(as.getLastRow()+1, 1));
        fj.deleteRow(a);
        a = a-1;
        lr = fj.getLastRow();
    }
  }
}

function archive_by_updated(){
  var latest_update_ids = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var as = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive');
  var update_id_Column = getColumnFromName("Latest Update ID:",ss)
  var lr = ss.getLastRow();
  for (var a=1; a<=lr; a++){
    var id = ss.getRange(a, update_id_Column).getValue();
    if (id){
    latest_update_ids.push(id);
    }
  }
  var fj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flattened JSON');
  lr = fj.getLastRow();
  var lc = fj.getLastColumn();
  for (var a=2;a<=lr;a++){
    if (!latest_update_ids.includes(fj.getRange(a,3).getValue())){
        fj.getRange(a,1,1,lc).moveTo(as.getRange(as.getLastRow()+1, 1));
        fj.deleteRow(a);
        a = a-1;
        lr = fj.getLastRow();
    }
  }
}


function writeETAs(obj){
  var etaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ETA');
  var containerColumn = getColumnFromName('Container ID',etaSheet)
  var containers_list = etaSheet.getRange(1,containerColumn,etaSheet.getLastRow(),1).getValues().map(x => x[0])
  var etaData = checkETAorATA(obj)
  if (!containers_list.includes(obj.payload.container_id)){
    if (!etaData){
      etaData = ["",""]
    }
    if (obj.payload.destination_port.geolocation){
      if (obj.payload.destination_port.geolocation.latitude && obj.payload.destination_port.geolocation.longitude){
        var geo = obj.payload.destination_port.geolocation.latitude + ", " + obj.payload.destination_port.geolocation.longitude
      }
      else geo = ""
    }
    else {
      var geo = ""
    }

    etaSheet.appendRow([
      obj.payload.container_id,
      obj.payload.carrier_scac,
      obj.payload.destination_port.unlocode,
      geo,
      etaData[0],
      etaData[2] ? etaData[1] : "", // Write to "Initial ETA" column if planned is true
      etaData[2] ? etaData[1] : "", // Write to "Current ETA" column if planned is true
      etaData[2] ? "" : etaData[1], // Write to "ATA" column if planned is false
    ]);
  } else if (etaData) {
    var row = containers_list.indexOf(obj.payload.container_id) + 1;

    if (etaData[2]){
      currentETAColumn = getColumnFromName('Current ETA', etaSheet);
      //currentETAValue = etaSheet.getRange(row, currentETAColumn).getValue().valueOf();
      ataColumn = getColumnFromName('ATA', etaSheet);
      //ataValue = etaSheet.getRange(row, ataColumn).getValue().valueOf();
  
      //if(ataValue != null && currentETAValue.valueOf() > ataValue.valueOf()) {
      etaSheet.getRange(row, currentETAColumn).clearContent();
      etaSheet.getRange(row, ataColumn).clearContent();
      //}
    }
    var etaOrATAColumn = etaData[2] ? getColumnFromName('Current ETA', etaSheet) : getColumnFromName('ATA', etaSheet);
    etaSheet.getRange(row, etaOrATAColumn).setValue(etaData[1]);
  }
}

function delete_by_subscribed(){
  var activeRefs = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Containers');
  var reference_id_Column = getColumnFromName("Reference ID",ss)
  var lr = ss.getLastRow();
  for (var a=1; a<=lr; a++){
    var ref = ss.getRange(a, reference_id_Column).getValue();
    if (ref){
    activeRefs.push(ref);
    }
  }
  console.log(activeRefs)
  var as = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive');
  lr = as.getLastRow();
  console.log(lr)
  for (var a=lr;a>=2;a--){
    if (!activeRefs.includes(as.getRange(a,3).getValue()) && !activeRefs.includes(as.getRange(a,2).getValue())){
        console.log(a)
        as.deleteRow(a);
    }
  }
}
