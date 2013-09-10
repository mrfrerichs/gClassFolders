function setResumeTrigger(lock) {
  lock.releaseLock();
  ScriptApp.newTrigger('createClassFolders').timeBased().after(30000).create();
  Browser.msgBox("Folder creation process will resume in 30 sec to avoid timeout");
  return;
}

function removeResumeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='createClassFolders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  return;
}


//Because we can't know whether a user is installing gClassFolders in a spreadsheet with other sheets
//we need a way to access the roster from the sheet's immutable ID
//This function is to be used whenever we need to get the roster sheet
// (note: Sheet ID and sheet index are different.  Index is mutable, id is not)
function getRosterSheet() {
  var sheetId = parseInt(ScriptProperties.getProperty('sheetId'));
  if (this.SSKEY) {
    var ss = SpreadsheetApp.openById(this.SSKEY);
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheets = ss.getSheets();
  var tempSheetId = '';
  for (var i=0; i<sheets.length; i++) {
    tempSheetId = sheets[i].getSheetId();
    if (tempSheetId==sheetId) {
      return sheets[i];
    } 
  }
  return;
}


//used to sort roster sheet by classname, period, and last name
function sortsheet(classIndex, perIndex, lNameIndex) {
  var sheet = getRosterSheet();
  if ((!classIndex)||(!perIndex)||(lNameIndex)) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    classIndex = headers.indexOf(this.labels().class + t(" Name"));
    perIndex = headers.indexOf(this.labels().period + " ~" + t("Optional") + "~");
    lNameIndex = headers.indexOf(t("Student Last Name"));
  }
  try {
    sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).sort([classIndex+1, perIndex+1, lNameIndex+1]);
  } catch(err) {
    Browser.msgBox(t("You cannot sort until you have entered student class enrollments"))
  }
  //sort by cls then by Per
  //Logger.log(sheet);
}


function getClassRowNumsFromRSF(dataRange, indices, rsfId) {
  var classRowNums = [];
    for (var i=1; i<dataRange.length; i++) {
      if (dataRange[i][indices.rsfIdIndex]==rsfId) {
        classRowNums.push(i+1);
      }
    }
  return classRowNums;
}


function getClassRowNumsFromCRF(dataRange, indices, crfId) {
  var classRowNums = [];
    for (var i=1; i<dataRange.length; i++) {
      if (dataRange[i][indices.crfIdIndex]==crfId) {
        classRowNums.push(i+1);
      }
    }
  return classRowNums;
}

function getSectionsNotSelected(dataRange, indices, rsfIds, crfId) {
  var rootsNotSelected = [];
  for (var i=1; i<dataRange.length; i++ ) {
    if ((rsfIds.indexOf(dataRange[i][indices.rsfIdIndex])==-1)&&(dataRange[i][indices.crfIdIndex]==crfId)&&(rootsNotSelected.indexOf(dataRange[i][indices.rsfIdIndex])==-1)) {
      rootsNotSelected.push(dataRange[i][indices.rsfIdIndex]);
    }
  }
  return rootsNotSelected;
}


function getClassRoster(dataRange, indices, className, per) {
  var crfId = ''; // class root folder id
  var rsfId = ''; // root student folder id
  var cefId = ''; // class edit folder id
  var cvfId = ''; // class view folder id
  var classRows = [];
  classRows.push(dataRange[0])
  for (var i=1; i<dataRange.length; i++) {
    if (per) {
      if((dataRange[i][indices.clsNameIndex]==className)&&(dataRange[i][indices.clsPerIndex]==per)) {
        classRows.push(dataRange[i]);
      }
    }
  if ((!per)||(per=='')) {
      if(dataRange[i][indices.clsNameIndex]==className) {
        classRows.push(dataRange[i]);
      }
    }
  }
  return classRows;
}


function getClassRosterAsObjects(dataRange, indices, className, per) {
  var crfId = ''; // class root folder id
  var rsfId = ''; // root student folder id
  var cefId = ''; // class edit folder id
  var cvfId = ''; // class view folder id
  var classRows = [];
  var rowNums = [];
  for (var i=1; i<dataRange.length; i++) {
    if (per) {
      if((dataRange[i][indices.clsNameIndex]==className)&&(dataRange[i][indices.clsPerIndex]==per)) {
        classRows.push(dataRange[i]);
        rowNums.push(i+1);
      }
    }
    if ((!per)||(per=='')) {
      if(dataRange[i][indices.clsNameIndex]==className) {
        classRows.push(dataRange[i]);
        rowNums.push(i+1);
      }
    }
  }
  var studentObjects = [];
  for (var i=0; i<classRows.length; i++) {
    studentObjects[i] = new Object();
    studentObjects[i]['sFName'] = classRows[i][indices.sFnameIndex];
    studentObjects[i]['sLName'] = classRows[i][indices.sLnameIndex];
    studentObjects[i]['sEmail'] = classRows[i][indices.sEmailIndex];
    studentObjects[i]['dbfId'] = classRows[i][indices.dbfIdIndex]; 
    studentObjects[i]['cvfId'] = classRows[i][indices.cvfIdIndex]; 
    studentObjects[i]['cefId'] = classRows[i][indices.cefIdIndex];
    studentObjects[i]['rsfId'] = classRows[i][indices.rsfIdIndex];
    studentObjects[i]['crfId'] = classRows[i][indices.crfIdIndex];
    studentObjects[i]['tfId'] = classRows[i][indices.tfIdIndex];
    if (indices.scfIdIndex!=-1) {
      studentObjects[i]['scfId'] = classRows[i][indices.scfIdIndex];
    }
    studentObjects[i]['clsName'] = classRows[i][indices.clsNameIndex];
    studentObjects[i]['clsPer'] = classRows[i][indices.clsPerIndex];
    studentObjects[i]['tEmail'] = classRows[i][indices.tEmailIndex];
    studentObjects[i]['row'] = rowNums[i];
  }
  return studentObjects;
}




function getClassFolderId(classRoster, folderIndex) {
  var folderId="";
  for (var i=1; i<classRoster.length; i++) {
    if (classRoster[i][folderIndex]!="") {
      folderId = classRoster[i][folderIndex];
      return folderId;
    }
  }
  return folderId;
}


function getUniqueClassNames(dataRange, clsNameIndex, crfIdIndex) {
  var classNames = [];
  for (var i=1; i<dataRange.length; i++) {
    var thisClassName = dataRange[i][clsNameIndex];
    var thisClassRoot = dataRange[i][crfIdIndex];
    if ((classNames.indexOf(thisClassName)==-1)&&(thisClassName!='')&&(thisClassRoot!='')) {
      classNames.push(thisClassName);
    }
  }
  classNames.sort();
  return classNames;
}


function returnEmailAsArray(emailValue) {
  emailValue = emailValue.replace(/\s+/g, '');
  var emailArray = emailValue.split(",");
  return emailArray;  
}


function getUniqueClassPeriods(dataRange, clsNameIndex, clsPerIndex, rsfIdIndex, labelObject) {
  var classPers = [];
  for (var i=0; i<dataRange.length; i++) {
    var thisClassPer = dataRange[i][clsNameIndex] + " " + labelObject.period + " " + dataRange[i][clsPerIndex];
    var thisStudentRoot = dataRange[i][rsfIdIndex];
    if ((classPers.indexOf(thisClassPer)==-1)&&(thisClassPer!='')&&(thisStudentRoot!='')) {
      classPers.push(thisClassPer);
    }
  }
  classPers.sort();
  return classPers;
}

function getRootClassFoldersByRSF(dataRange, rsfId, rsfIdIndex, crfIdIndex, cefIdIndex, cvfIdIndex) {
  for (var i=1; i<dataRange.length; i++) {
    var thisRsfId = dataRange[i][rsfIdIndex];
    if (thisRsfId == rsfId) {
      var obj = new Object();
      obj.crfId = dataRange[i][crfIdIndex];
      obj.cefId = dataRange[i][cefIdIndex];
      obj.cvfId = dataRange[i][cvfIdIndex];
      return obj;
    }
  }
  return;
}

function getTeacherEmailsByRSF(dataRange, rsfId, rsfIdIndex, tEmailIndex) {
  for (var i=1; i<dataRange.length; i++) {
    var thisRsfId = dataRange[i][rsfIdIndex];
    if (thisRsfId == rsfId) {
      var obj = new Object();
      obj.tEmails = dataRange[i][tEmailIndex];
      return obj;
    }
  }
  return;
}



function getUniqueClassPeriodObjects(dataRange, clsNameIndex, clsPerIndex, rsfIdIndex, labelObj) {
  var classPers = [];
  var processed = [];
  var k = 0;
  for (var i=1; i<dataRange.length; i++) {
    var thisClassPer = dataRange[i][clsNameIndex];
    if (dataRange[i][clsPerIndex]!='') {
      thisClassPer += " " + labelObj.period + " " + dataRange[i][clsPerIndex];
    }
    var thisStudentRoot = dataRange[i][rsfIdIndex];
    if ((processed.indexOf(thisClassPer)==-1)&&(thisClassPer!='')&&(thisStudentRoot!='')) {
      classPers[k] = new Object();
      classPers[k].classPer = thisClassPer;
      classPers[k].rsfId = thisStudentRoot;
      processed.push(thisClassPer);
      k++;
    }
  }
  classPers.sort(
    function compareNames(a, b) {
      var nameA = a.classPer.toLowerCase( );
      var nameB = b.classPer.toLowerCase( );
      if (nameA < nameB) {return -1}
      if (nameA > nameB) {return 1}
      return 0;
    })
    return classPers;
}

function saveIndices(indices) {
  if (indices) {
    var indicesString = Utilities.jsonStringify(indices);
    ScriptProperties.setProperty('indices', indicesString);
  }
}


function writeProperties() {
  var properties = ScriptProperties.getProperties();
  var propertyArray = [];
  var i = 0;
  for (var key in properties) {
    propertyArray[i] = [];
    propertyArray[i][0] = key;
    propertyArray[i][1] = properties[key]
    i++;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName('Properties');
  if (!sheet) {
    sheet = ss.insertSheet('Properties');
  }
  sheet.getRange(1, 1, propertyArray.length, 2).setValues(propertyArray);
  sheet.getRange("A1").setComment("This sheet is used by gClassHub to understand how your roster is organized.");
}


function returnIndices(dataRange, labelObject) {
  var sheet = getRosterSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var sFnameIndex = headers.indexOf(t('Student First Name'));
  if (sFnameIndex==-1) {
    badHeaders();
    return;
  }
  var sLnameIndex = headers.indexOf(t('Student Last Name'));
  if (sLnameIndex==-1) {
    badHeaders();
    return;
  }
  var sEmailIndex = headers.indexOf(t('Student Email'));
  if (sEmailIndex==-1) {
    badHeaders();
    return;
  }
  var clsNameIndex  = headers.indexOf(labelObject.class +t(' Name'));
  if (clsNameIndex==-1) {
    badHeaders();
    return;
  }
  var clsPerIndex = headers.indexOf(labelObject.period + " ~" + t('Optional') + "~");
  if (clsNameIndex==-1) {
    badHeaders();
    return;
  }
  var tEmailIndex = headers.indexOf(t('Teacher Email(s)'));
  if (tEmailIndex==-1) {
    badHeaders();
    return;
  }  
  
  //Add columns for tracking status of folder creation and share if they don't already exist
  //retrieve their indices
  
  var sDropStatusIndex = headers.indexOf(t("Status: Student " + labelObject.dropBox));
  if (sDropStatusIndex==-1) {
    sheet.getRange(1,lastCol+1,1,2).setValues([[t("Status: Student " + labelObject.dropBox),t("Status: Teacher Share")]]).setComment(t("Don't change this header. When gClassFolders is run, class and dropbox folders will be created or updated for any students without a value in this column. To update a student's email address or name, just clear their status value and run again. To move students between classes, use the menu."));
    headers.push(t("Status: Student ") + labelObject.dropbox);
    headers.push(t("Status: Teacher Share"));
    SpreadsheetApp.flush()
  }
  sDropStatusIndex = headers.indexOf(t("Status: Student " + labelObject.dropBox));
  var tShareStatusIndex = headers.indexOf(t("Status: Teacher Share"));
  var dbfIdIndex = headers.indexOf(t('Student') + ' ' + labelObject.dropBox + ' Id');
  if (dbfIdIndex==-1) {
    createFolderIdHeadings(); //create Folder ID headings if they don't exist
  }
  
  //refresh headers to pull in new folder id columns;
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  //Get indices of folder Id columns
  var indices = new Object();
  indices.dbfIdIndex = headers.indexOf(t('Student') + " " + labelObject.dropBox + ' Id');
  indices.crfIdIndex = headers.indexOf(t('Class Root Folder') + ' Id');
  indices.cvfIdIndex = headers.indexOf(t('Class View Folder') + ' Id');
  indices.cefIdIndex = headers.indexOf(t('Class Edit Folder') + ' Id');
  indices.rsfIdIndex = headers.indexOf(t('Root Student Folder') + ' Id');
  indices.tfIdIndex = headers.indexOf(t('Teacher Folder') + ' Id');
  indices.scfIdIndex = headers.indexOf(t('Student') + " " + labelObject.class + " " + t('Folder') + ' Id');
  indices.sFnameIndex = sFnameIndex;
  indices.sLnameIndex = sLnameIndex;
  indices.sEmailIndex = sEmailIndex;
  indices.clsNameIndex = clsNameIndex;
  indices.clsPerIndex = clsPerIndex;
  indices.tEmailIndex = tEmailIndex;
  indices.sDropStatusIndex = sDropStatusIndex;
  indices.tShareStatusIndex = tShareStatusIndex;
  indices.dbfIdIndex = dbfIdIndex;
  return indices;
}




function gClassFolders_institutionalTrackingUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (!(institutionalTrackingString)) {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
  }
  var app = UiApp.createApplication().setTitle('Hello there! Help us track the usage of this script').setHeight(400);
  if ((!(institutionalTrackingString))||(!(eduSetting))) {
    var helptext = app.createLabel(t("You are most likely seeing this prompt because this is the first time you are using a Google Apps script created by New Visions for Public Schools, 501(c)3. If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering tracking information here will save it to your user credentials and enable tracking for any other New Visions scripts that use this feature. No personal info will ever be collected.")).setStyleAttribute('marginBottom', '10px');
  } else {
  var helptext = app.createLabel(t("If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering or modifying tracking information here will save it to your user credentials and enable tracking for any other scripts produced by New Visions for Public Schools, 501(c)3, that use this feature. No personal info will ever be collected.")).setStyleAttribute('marginBottom', '10px');
  }
  var panel = app.createVerticalPanel();
  var gridPanel = app.createVerticalPanel().setId("gridPanel").setVisible(false);
  var grid = app.createGrid(4,2).setId('trackingGrid').setStyleAttribute('background', 'whiteSmoke').setStyleAttribute('marginTop', '10px');
  var checkHandler = app.createServerHandler('gClassFolders_refreshTrackingGrid').addCallbackElement(panel);
  var checkBox = app.createCheckBox(t('Participate in institutional usage tracking.  (Only choose this option if you know your institution\'s Google Analytics tracker Id.)')).setName('trackerSetting').addValueChangeHandler(checkHandler);  
  var checkBox2 = app.createCheckBox(t('Let') + " New Visions for Public Schools, 501(c)3" + " " + t('and') +  " gClassFolders " + t('co-author') + " Bjorn Behrendt " + t("know you're an educational user.")).setName('eduSetting');  
  if ((institutionalTrackingString == "not participating")||(institutionalTrackingString=='')) {
    checkBox.setValue(false);
  } 
  if (eduSetting=="true") {
    checkBox2.setValue(true);
  }
  var institutionNameFields = [];
  var trackerIdFields = [];
  var institutionNameLabel = app.createLabel(t('Institution Name'));
  var trackerIdLabel = app.createLabel(t('Google Analytics Tracker Id')+' (UA-########-#)');
  grid.setWidget(0, 0, institutionNameLabel);
  grid.setWidget(0, 1, trackerIdLabel);
  if ((institutionalTrackingString)&&((institutionalTrackingString!='not participating')||(institutionalTrackingString==''))) {
    checkBox.setValue(true);
    gridPanel.setVisible(true);
    var institutionalTrackingObject = Utilities.jsonParse(institutionalTrackingString);
  } else {
    var institutionalTrackingObject = new Object();
  }
  for (var i=1; i<4; i++) {
    institutionNameFields[i] = app.createTextBox().setName('institution-'+i);
    trackerIdFields[i] = app.createTextBox().setName('trackerId-'+i);
    if (institutionalTrackingObject) {
      if (institutionalTrackingObject['institution-'+i]) {
        institutionNameFields[i].setValue(institutionalTrackingObject['institution-'+i]['name']);
        if (institutionalTrackingObject['institution-'+i]['trackerId']) {
          trackerIdFields[i].setValue(institutionalTrackingObject['institution-'+i]['trackerId']);
        }
      }
    }
    grid.setWidget(i, 0, institutionNameFields[i]);
    grid.setWidget(i, 1, trackerIdFields[i]);
  } 
  var help = app.createLabel(t("Enter up to three institutions, with Google Analytics tracker Id's.")).setStyleAttribute('marginBottom','5px').setStyleAttribute('marginTop','10px');
  gridPanel.add(help);
  gridPanel.add(grid); 
  panel.add(helptext);
  panel.add(checkBox2);
  panel.add(checkBox);
  panel.add(gridPanel);
  var button = app.createButton(t("Save settings"));
  var saveHandler = app.createServerHandler('gClassFolders_saveInstitutionalTrackingInfo').addCallbackElement(panel);
  button.addClickHandler(saveHandler);
  panel.add(button);
  app.add(panel);
  ss.show(app);
  return app;
}

function gClassFolders_refreshTrackingGrid(e) {
  var app = UiApp.getActiveApplication();
  var gridPanel = app.getElementById("gridPanel");
  var grid = app.getElementById("trackingGrid");
  var setting = e.parameter.trackerSetting;
  if (setting=="true") {
    gridPanel.setVisible(true);
  } else {
    gridPanel.setVisible(false);
  }
  return app;
}

function gClassFolders_saveInstitutionalTrackingInfo(e) {
  var app = UiApp.getActiveApplication();
  var eduSetting = e.parameter.eduSetting;
  var oldEduSetting = UserProperties.getProperty('eduSetting')
  if (eduSetting == "true") {
    UserProperties.setProperty('eduSetting', 'true');
  }
  if ((oldEduSetting)&&(eduSetting=="false")) {
    UserProperties.setProperty('eduSetting', 'false');
  }
  var trackerSetting = e.parameter.trackerSetting;
  if (trackerSetting == "false") {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
    app.close();
    return app;
  } else {
    var institutionalTrackingObject = new Object;
    for (var i=1; i<4; i++) {
      var checkVal = e.parameter['institution-'+i];
      if (checkVal!='') {
        institutionalTrackingObject['institution-'+i] = new Object();
        institutionalTrackingObject['institution-'+i]['name'] = e.parameter['institution-'+i];
        institutionalTrackingObject['institution-'+i]['trackerId'] = e.parameter['trackerId-'+i];
        if (!(e.parameter['trackerId-'+i])) {
          Browser.msgBox(t("You entered an institution without a Google Analytics Tracker Id"));
          gClassFolders_institutionalTrackingUi()
        }
      }
    }
    var institutionalTrackingString = Utilities.jsonStringify(institutionalTrackingObject);
    UserProperties.setProperty('institutionalTrackingString', institutionalTrackingString);
    onOpen();
    Browser.msgBox(t("Once you have added your course rosters, use the gClassFolders menu to generate class folders."));
    app.close();
    return app;
  }
}



function gClassFolders_createInstitutionalTrackingUrls(institutionTrackingObject, encoded_page_name, encoded_script_name) {
  for (var key in institutionTrackingObject) {
   var utmcc = gClassFolders_createGACookie();
  if (utmcc == null)
    {
      return null;
    }
  var encoded_page_name = encoded_script_name+"/"+encoded_page_name;
  var trackingId = institutionTrackingObject[key].trackerId;
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.gClassFolders-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  if (ga_url_full)
    {
      var response = UrlFetchApp.fetch(ga_url_full);
    }
  }
}



function gClassFolders_createGATrackingUrl(encoded_page_name)
{
  var utmcc = gClassFolders_createGACookie();
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (eduSetting=="true") {
    encoded_page_name = "edu/" + encoded_page_name;
  }
  if (utmcc == null)
    {
      return null;
    }
 
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.gClassFolders-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-38070753-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  return ga_url_full;
}


function gClassFolders_createGACookie()
{
  var a = "";
  var b = "100000000";
  var c = "200000000";
  var d = "";

  var dt = new Date();
  var ms = dt.getTime();
  var ms_str = ms.toString();
 
  var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
  var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
  if ((gClassFolders_teacher_uid == null) && (gClassFolders_school_uid == ""))
    {
      // shouldn't happen unless user explicitly removed flubaroo_uid from properties.
      return null;
    }
  
  if (gClassFolders_teacher_uid) {
    a = gClassFolders_teacher_uid.substring(0,9);
    d = gClassFolders_teacher_uid.substring(9);
  }
  
  if (gClassFolders_school_uid) {
    a = gClassFolders_school_uid.substring(0,9);
    d = gClassFolders_school_uid.substring(9);
  }
  
  utmcc = "__utma%3D451096098." + a + "." + b + "." + c + "." + d 
          + ".1%3B%2B__utmz%3D451096098." + d + ".1.1.utmcsr%3D(direct)%7Cutmccn%3D(direct)%7Cutmcmd%3D(none)%3B";
 
  return utmcc;
}



function gClassFolders_logStudentFolderCreation()
{
  var ga_url = gClassFolders_createGATrackingUrl("Student%20Class%20Folder%20Created");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"Student%20Class%20Folder%20Created", "gClassFolders");
  }
}


function gClassFolders_logTeacherClassFolderCreated()
{
  var ga_url = gClassFolders_createGATrackingUrl("Teacher%20Class%20Folder%20Created");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"Teacher%20Class%20Folder%20Created", "gClassFolders");
  }
}


function gClassFolders_logStudentClassFolderArchived()
{
  var ga_url = gClassFolders_createGATrackingUrl("Student%20Class%20Folder%20Archived");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"Student%20Class%20Folder%20Archived", "gClassFolders");
  }
}


function gClassFolders_getInstitutionalTrackerObject() {
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  if ((institutionalTrackingString)&&(institutionalTrackingString != "not participating")) {
    var institutionTrackingObject = Utilities.jsonParse(institutionalTrackingString);
    return institutionTrackingObject;
  }
  if (!(institutionalTrackingString)||(institutionalTrackingString=='')) {
    gClassFolders_institutionalTrackingUi();
    return;
  }
}


function gClassFolders_logRepeatTeacherInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("Repeat%20Teacher%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
      var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"Repeat%20Teacher%20Install", "gClassFolders");
  }
}



function gClassFolders_logRepeatSchoolInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("Repeat%20School%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
      var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"Repeat%20School%20Install", "gClassFolders");
  }
}


function gClassFolders_logFirstTeacherInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("First%20Teacher%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
   if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"First%20Teacher%20Install", "gClassFolders");
  }
}


function gClassFolders_logFirstSchoolInstall()
{
  var ga_url = gClassFolders_createGATrackingUrl("First%20School%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = gClassFolders_getInstitutionalTrackerObject();
   if (institutionalTrackingObject) {
    gClassFolders_createInstitutionalTrackingUrls(institutionalTrackingObject,"First%20School%20Install", "gClassFolders");
  }
}



function setgClassFoldersTeacherUid()
{ 
  var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
  if (gClassFolders_teacher_uid == null || gClassFolders_teacher_uid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty("gClassFolders_teacher_uid", ms_str);
      gClassFolders_logFirstTeacherInstall();
    }
}


function setgClassFoldersSchoolUid()
{ 
  var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
  if (gClassFolders_school_uid == null || gClassFolders_school_uid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty("gClassFolders_school_uid", ms_str);
      gClassFolders_logFirstSchoolInstall();
    }
}


function setgClassFoldersTeacherSid()
{ 
  var gClassFolders_teacher_sid = ScriptProperties.getProperty("gClassFolders_teacher_sid");
  if (gClassFolders_teacher_sid == null || gClassFolders_teacher_sid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      ScriptProperties.setProperty("gClassFolders_teacher_sid", ms_str);
      var gClassFolders_teacher_uid = UserProperties.getProperty("gClassFolders_teacher_uid");
      if (gClassFolders_teacher_uid != null || gClassFolders_teacher_uid != "") {
        gClassFolders_logRepeatTeacherInstall();
      }
    }
}


function setgClassFoldersSchoolSid()
{ 
  var gClassFolders_teacher_sid = ScriptProperties.getProperty("gClassFolders_school_sid");
  if (gClassFolders_teacher_sid == null || gClassFolders_teacher_sid == "")
    {
      // user has never installed gClassFolders before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      ScriptProperties.setProperty("gClassFolders_school_sid", ms_str);
      var gClassFolders_school_uid = UserProperties.getProperty("gClassFolders_school_uid");
      if (gClassFolders_school_uid != null || gClassFolders_school_uid != "") {
        gClassFolders_logRepeatSchoolInstall();
      }
    }
}
