// gClassFolders
// Original resource by EdListen.com 
// Original Author: Bjorn Behrendt bj@edlisten.com
// Version 2.1.2-dev (9/9/2013) a collaboration with Andrew Stillman astillman@gmail.com and YouPD.org, a project of New Visions for Public Schools.
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html


var GCLASSICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var GCLASSLAUNCHERICONURL = 'https://sites.google.com/site/gclassfolders/_/rsrc/1360538205205/config/customLogo.gif?revision=1';
var userEmail = Session.getEffectiveUser().getEmail();
var SSKEY = ScriptProperties.getProperty('ssKey');

// This list was taken from the list of available languages in Google Translate service, responsible for our UI internationalization.
var googleLangList = ['English: en','Afrikaans: af','Albanian: sq','Arabic: ar','Azerbaijani: az','Basque: eu','Bengali: bn','Belarusian: be','Bulgarian: bg','Catalan: ca','Chinese Simplified: zh-CN','Chinese Traditional: zh-TW','Croatian: hr','Czech: cs','Danish: da','Dutch: nl','Esperanto: eo','Estonian: et','Filipino: tl','Finnish: fi','French: fr','Galician: gl','Georgian: ka','German: de','Greek: el','Gujarati: gu','Haitian Creole: ht','Hebrew: iw','Hindi: hi','Hungarian: hu','Icelandic: is','Indonesian: id','Irish: ga','Italian: it','Japanese: ja','Kannada: kn','Korean: ko','Latin: la','Latvian: lv','Lithuanian: lt','Macedonian: mk','Malay: ms','Maltese: mt','Norwegian: no','Persian: fa','Polish: pl','Portuguese: pt','Romanian: ro','Russian: ru','Serbian: sr','Slovak: sk','Slovenian: sl','Spanish: es','Swahili: sw','Swedish: sv','Tamil: ta','Telugu: te','Thai: th','Turkish: tr','Ukrainian: uk','Urdu: ur','Vietnamese: vi','Welsh: cy','Yiddish: yi'];

// This object is responsible for returning the values for custom labels for "Assignment Folder", "Class", and "Period" throughout the script
var labels = function() { var labels = CacheService.getPublicCache().get('labels');
                         if (!labels) {
                           labels = ScriptProperties.getProperty('labels');
                           CacheService.getPublicCache().put('labels', labels, 660);
                         }
                         if (labels) {
                           labels = Utilities.jsonParse(labels);
                         } else {
                           labels =  {dropBox: "Assignment Folder", dropBoxes: "Assignment Folders", class: "Class", classes: "Classes", period: "Period"};
                         }
                         return labels;
                        }

//This function only runs once, when the gallery installation process occurs.  
// Helps orient new users to the script and forces the menu to appear immediately upon installation.
function onInstall() {
  var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  CacheService.getPublicCache().put('lang', locale.substring(0,2), 6000);
  onOpen();
}



//This function runs automatically when the spreadsheet opens, and provides the initial menu to the script.
//Defined separately from myOnOpen() to avoid issues with ScriptProperties not being able to be called from a built in trigger
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries[0] =  {name: "What is gClassFolders?", functionName:"gClassFolders_whatIs"};
  menuEntries[1] = {name: "Initial settings", functionName:"gClassFolders_lang"};
  ss.addMenu("gClassFolders", menuEntries);
  myOnOpen(menuEntries);
}





//Secondary menu function built to check for initial setup state and repopulate menu
//with appropriate choices
function myOnOpen(menuEntries) {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lang = ScriptProperties.getProperty('lang');
  var mode = ScriptProperties.getProperty('mode');
  menuEntries = [];
  menuEntries[0] = {name: t("What is gClassFolders?"), functionName:"gClassFolders_whatIs"};
  if ((!lang)||(!mode)) {
    menuEntries.push({name: "Initial settings", functionName:"gClassFolders_lang"});
  }
  var trackerSettings = UserProperties.getProperty('institutionalTrackingString');
  if ((lang)&&(mode)&&(!trackerSettings)) {
    menuEntries.push({name: t("Help us track usage"),functionName: "gClassFolders_institutionalTrackingUi"});
  }
  //Check to see whether user has created roster sheet or the roster sheet 
  //has been deleted.  If not, prompt user to create one.
  var sheetId = ScriptProperties.getProperty('sheetId');
  if ((!sheetId)&&(lang)&&(mode)) {
    menuEntries.push({name: t("Set / create roster sheet"), functionName: "createRosterSheet"});
  } else if ((sheetId)&&(lang)&&(mode)) {  //otherwise prompt for next steps
    menuEntries.push(null); // line separator
    menuEntries.push({name: t("Sort sheet by") + this.labels().class + ", " + this.labels().period + ", " + t('last name'), functionName: "sortsheet"});
    menuEntries.push({name: t("Create new folders and shares"), functionName: "createClassFolders"});
    menuEntries.push({name: t("Perform bulk operations on selected student(s)"), functionName: "bulkOperationsUi"});
  } 
  //Check to see whether roster exists and folder creation has already taken place
  var alreadyRan = ScriptProperties.getProperty('alreadyRan');
  if ((sheetId)&&(alreadyRan)) {
    menuEntries.push({name: t("Get gClassHub URL"), functionName:"getGClassHubUrl"});
  }
  ss.updateMenu("gClassFolders", menuEntries)
}





//Builds the UI for initial settings
function gClassFolders_lang() {
  var app = UiApp.createApplication().setHeight(550);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var title = app.createLabel(t("Choose your preferred interface language.")).setId('title').setStyleAttribute('fontSize', '18px').setStyleAttribute('marginBottom', '8px');
  var label = app.createLabel(t("Note: This will affect all column heading and menu options. Once this setting is chosen, it cannot be changed.")).setId('langLabel');
  var listBox = app.createListBox().setName('lang');
  var langArray = this.googleLangList;
  var langName = '';
  var langCode = ''; 
  for (var i=0; i<langArray.length; i++) {
    langName = langArray[i].split(": ")[0];
    langCode = langArray[i].split(": ")[1];
    listBox.addItem(langName, langCode);
  }
  var saveHandler = app.createServerHandler('saveLangSettings').addCallbackElement(panel);
  var refreshHandler = app.createServerHandler('refreshLangPanel').addCallbackElement(panel);
  listBox.addChangeHandler(refreshHandler);
  var defaultLabels = {dropBox: "Assignment Folder", dropBoxes: "Assignment Folders", class: "Class", classes: "Classes", period: "Period"};
  var namingLabel = app.createLabel(t("The terms below can be renamed, and require custom translation. These labels determine how important folders and columns will be named by gClassFolders.")).setId('namingLabel').setStyleAttribute('marginTop', '5px');
  var namingGrid = app.createGrid(5, 2).setId('namingGrid').setCellPadding(3);
  namingGrid.setWidget(0, 0, app.createLabel(defaultLabels.dropBox)).setWidget(0, 1, app.createTextBox().setName('dropBox').setValue(t(defaultLabels.dropBox)))
  .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes)).setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(t(defaultLabels.dropBoxes)))
  .setWidget(2, 0, app.createLabel(defaultLabels.class)).setWidget(2, 1, app.createTextBox().setName('class').setValue(t(defaultLabels.class)))
  .setWidget(3, 0, app.createLabel(defaultLabels.classes)).setWidget(3, 1, app.createTextBox().setName('classes').setValue(t(defaultLabels.classes)))
  .setWidget(4, 0, app.createLabel(defaultLabels.period)).setWidget(4, 1, app.createTextBox().setName('period').setValue(t(defaultLabels.period)));
  panel.add(title);
  panel.add(label);
  panel.add(listBox);
  panel.add(namingLabel);
  panel.add(namingGrid);
  
  var title2 = app.createLabel(t("Indicate how you plan to run gClassFolders")).setId('title2').setStyleAttribute('fontSize', '18px').setStyleAttribute('marginTop', '20px');
  var listBox2 = app.createListBox().setName('mode').setStyleAttribute('marginTop', '10px').setId('listBox2');
  listBox2.addItem(t('Single Teacher Mode'), 'teacher');
  listBox2.addItem(t('School Mode'), 'school');
  var modeChangeHandler = app.createServerHandler('refreshModeDescription').addCallbackElement(panel);
  listBox2.addChangeHandler(modeChangeHandler);
  var descriptionText = t('Meant for one teacher running gClassFolders from their own account. Single Teacher Mode is simpler, but offers fewer options for managing and archiving student work over time and across a student\'s classes if they have multiple teachers.');
  var description = app.createLabel(descriptionText).setStyleAttribute('backgroundColor', '#33FFFF').setStyleAttribute('padding', '5px').setStyleAttribute('margin', '15px').setId('description');
  var warningLabel = app.createLabel(t('Choose carefully! Once student folders have been generated, it is not possible to switch modes')).setId('warningLabel');
  panel.add(title2);
  panel.add(listBox2);
  panel.add(description);
  panel.add(warningLabel);  
  
  app.add(panel);
  app.add(app.createButton(t("Save"), saveHandler).setId('button'));
  ss.show(app);
  return app;
}

// Refreshes all UI elements in the initial settings panel to ensure that language changes
// render in panel as user changes language
// Note: technically the language setting gets stored on each refresh of the language selector.
// An alternate method using CacheService was attempted but the service appears to be broken in the Ui Context
// Worth a 2nd look in the future
function refreshLangPanel(e) {
  var app = UiApp.getActiveApplication();
  var title = app.getElementById('title');
  var label = app.getElementById('langLabel');
  var namingLabel = app.getElementById('namingLabel');
  var namingGrid = app.getElementById('namingGrid');
  var button = app.getElementById('button');
  var defaultLabels = {dropBox: "Assignment Folder", dropBoxes: "Assignment Folders", class: "Class", classes: "Classes", period: "Period"};
  var lang = e.parameter.lang;
  ScriptProperties.setProperty('lang', lang);
  //CacheService.getPublicCache().remove('lang');
  title.setText(t("Choose your preferred interface language."));
  label.setText(t("Note: This will affect all column headings and menu options. Once this setting is chosen, it cannot be changed."));
  namingLabel.setText(t("The terms below can be renamed, and require custom translation. These labels determine how important folders and columns will be named by gClassFolders."));
  namingGrid.setWidget(0, 0, app.createLabel(defaultLabels.dropBox)).setWidget(0, 1, app.createTextBox().setName('dropBox').setValue(t(defaultLabels.dropBox)))
  .setWidget(1, 0, app.createLabel(defaultLabels.dropBoxes)).setWidget(1, 1, app.createTextBox().setName('dropBoxes').setValue(t(defaultLabels.dropBoxes)))
  .setWidget(2, 0, app.createLabel(defaultLabels.class)).setWidget(2, 1, app.createTextBox().setName('class').setValue(t(defaultLabels.class)))
  .setWidget(3, 0, app.createLabel(defaultLabels.classes)).setWidget(3, 1, app.createTextBox().setName('classes').setValue(t(defaultLabels.classes)))
  .setWidget(4, 0, app.createLabel(defaultLabels.period)).setWidget(4, 1, app.createTextBox().setName('period').setValue(t(defaultLabels.period)));
  var title2 = app.getElementById('title2');
  var listBox2 = app.getElementById('listBox2');
  var description = app.getElementById('description');
  var warningLabel = app.getElementById('warningLabel');
  title2.setText(t("Indicate how you plan to run gClassFolders"));
  listBox2.clear();
  listBox2.addItem(t('Single Teacher Mode'), 'teacher');
  listBox2.addItem(t('School Mode'), 'school');
  var descriptionText = t('Meant for one teacher running gClassFolders from their own account. Single Teacher Mode is simpler, but offers fewer options for managing and archiving student work over time and across a student\'s classes if they have multiple teachers.');
  description.setText(descriptionText);
  warningLabel.setText(t('Choose carefully! Once student folders have been generated, it is not possible to switch modes.')).setId('warningLabel');
  button.setHTML(t("Save settings"));
  return app;
}





// Function to refresh just the mode description when dropdown value is changed.
function refreshModeDescription(e) {
  var app = UiApp.getActiveApplication();
  var mode = e.parameter.mode;
  var description = app.getElementById('description');
  var descriptionText = '';
  if (mode == 'teacher') {
    descriptionText = t('Meant for one teacher running gClassFolders from their own account. Single Teacher Mode is simpler, but offers fewer options for managing and archiving student work over time and across a student\'s classes if they have multiple teachers.'); 
    description.setStyleAttribute('backgroundColor', '#33FFFF');
  } else {
    descriptionText = t('School Mode is meant to be run from a "Role Account" -- a domain account dedicated to managing and archiving all student course work over multiple semesters or years. School mode is a better choice for multiple teachers, and helps organize all student work into career portfolios.');
    description.setStyleAttribute('backgroundColor', '#99CC99');
  }
  description.setText(descriptionText);
  return app;
}




// Saves language and mode settings
function saveLangSettings(e) {
  var properties = ScriptProperties.getProperties();
  var app = UiApp.getActiveApplication();
  var lang = e.parameter.lang;
  var dropBox= e.parameter.dropBox;
  var dropBoxes = e.parameter.dropBoxes;
  var class = e.parameter.class;
  var classes = e.parameter.classes;
  var period = e.parameter.period;
  var mode = e.parameter.mode;
  properties.labels = Utilities.jsonStringify({dropBox: dropBox, dropBoxes: dropBoxes, class: class, classes: classes, period: period});
  properties.lang = lang;
  properties.mode = mode;
  properties.ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptProperties.setProperties(properties);
  if ((mode=="school")&&(!properties.scriptUrl)) {
    howToPublishAsWebApp();
    app.close();
    return app;
  }
  createRosterSheet();
  app.close();
  return app;
}



// Ui providing step by step instructions for how to publish the script as a webApp
function howToPublishAsWebApp() {
  var app = UiApp.createApplication().setTitle('School mode requires you to publish gClassFolders as a web app').setHeight(540).setWidth(600);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var published = verifyPublished();
  var properties = ScriptProperties.getProperties();
  if (published) {
    createRosterSheet();
  } else {
  var panel = app.createVerticalPanel();
  var handler = app.createServerHandler('confirmSettings').addCallbackElement(panel);
  var button = app.createButton("Confirm my settings").addClickHandler(handler);
  var scrollpanel = app.createScrollPanel().setHeight("360px");
  var grid = app.createGrid(8, 2).setBorderWidth(0).setCellSpacing(0);
  var html = app.createHTML('gClassFolders SCHOOL MODE requires publishing this script as a web app, which will provide a URL that will be automatically sent to users to allow them to automatically add their active class folders to "My Drive." The instructions below explain how to publish this script as a web app.');
  panel.add(html);
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  panel.add(text1);
  grid.setWidget(0, 0, app.createHTML('1. Go to \'Tools->Script editor\' from the Spreadsheet that contains your form.</li>'));
  grid.setWidget(0, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/scriptEditor.jpg?attachauth=ANoY7cr_KgzyqGdhNrUJw2hAmcd7_PtO_uYYqXr32jY8Sd_hmyyhZsAaQP0wDTZkXjPtANdqOy8oCFGlXeic8g0gxWokWEuIRe_1xyE0LYzaXXxfdhG1BgElGKq4Lb3bgekcvdpgtsxo0NmTSjoebLJPpl-omhj9lSOaGYCZ9yKOsd0HUTbZYR8riA3KIbCkQ6Y71jlycopwz3PTGDJKOOJ1O-eiG4DDHv5J-gVFGolkjgV8cy0hDe49rXZ8zkMV8DvDgf1-b0bx&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, app.createHTML('2. Under the \'File\' menu in the Script Editor, select \'Manage versions\' and save a new version of the script. Because it\'s optional, you can leave the \'Describe what changed\' field blank.</li>'));
  grid.setWidget(2, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/mgversions.jpg?attachauth=ANoY7coyiztUURiQ3mEa8Rg05aw6uxuuMO0UXlLM8PK_xdGfJvp_Wz3S9Pq1JGmDV4Yfhyfks_z7vgs47FszLPJcYlLajcH1LvOQSASIp2vzvpIaUiPZz8fVdRxmv1IFqTuTf1AbjHaCNGUa3UhrBEHq-2qzJTJp3cK6K910C_L2rfFhdjuCA_z9OU6LMt69UKryskdzu5G-xl7bCWdNKFaXXmyBlxDlQMFDqQCW8oySbMqk-XBfJ3UyBe6WGQpohaIfcSFrCpoi&attredirects=0').setWidth("430px"));
  grid.setWidget(3, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/manage%20versions.png?attachauth=ANoY7cpVfU-WEwiPuOMKMIAzbK6EdA8xkmv_M2R8GKdlcGLC7mo00ZJykbBFrtJEZQHDpKVdvizQQnuyfGVc65iigmGuGr_ZwC2Z4rnh1V67_ogOJKXH2TWmDAafxa-q_5fngrasDYYN2w2-hR_eR95GoY6e5Rza-mtWb1iAp97Cm8n9kVHRk67dURdrdD5AIaS8ZOkse1MmfaN-ZJpMv7bLYBKpisq8GldTTjo7W55OUIJhFuDcxLEc__vguXArjfb9Pd_e2bZD&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(4, 0, "backgroundColor", "grey").setStyleAttribute(4, 1, "backgroundColor", "grey");
  grid.setWidget(5, 0, app.createHTML('3. Under the \'Publish\' menu in the Script Editor, select \'Deploy as web app\'. Choose the version you want to publish (usually #1). Under \'Execute web app as\', choose \'User accessing the web app\'. Under "Who has access to the app:", select "Anyone within mydomain.org."  This script will only ever reveal Documents in the course folders to which users already have access.'));
  grid.setWidget(5, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/deploy.jpg?attachauth=ANoY7cr12RO45sWJXv0gu9U9X1qSzMdu607iozuY20X9-iBtoVnQYaTLMIqwJWEWgPk7kU6M3XTxnOyQDrGA84LZI_Y4PT5QrOOcHgD6AxNChrDJI4OzaYwB_nYF0ylYeN6rtKPwSRvXYMyQy1-Gbkz5j9mZxneDkkuI12hTsMkdlzPNhnJz5EiJY32fn_s9mQ3_01X7_PL-L-66UfMi4Yb6XwK-XCUgzpHtQoh4PQl0X9oduYeCApo0buL0ftmQmC0sbkazjSxw&attredirects=0').setWidth("430px"));
  grid.setWidget(6, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/deploy2.jpg?attachauth=ANoY7cpvRpmwz3oTkN-V0TQTK-wNj0APSLlDB22dwLmDe-rHHLgvug-ZeTxZAWJcVAKxQQIxzA_ylvZulJ4KsQSYn9IYL7UME471t1QnQScnzjIpOydp5Qup-didvNU3sXcAggG8v0kuIO1oriOcf_tBZzzQCiFufvmiL-YKa9M5BRcT1N35eW8DCNhBVsvkHy_64yU85JynpJ7NMjiKVPPUJ4MtvnsNhlaV40k5TJHH3B-3QjOjLvyNRF0aeViPKPIjWThM1W13&attredirects=0').setWidth("430px"));
  grid.setWidget(7, 0, app.createHTML('4. Once you have published the script as a web app, you will be provided a URL. You don\'t need to do anything with this URL, as it will be sent automatically to users with custom arguments when their root folders are created.'));
  grid.setWidget(7, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/step4.jpg?attachauth=ANoY7cqIpA9FyKCBHy4ZJYj3jXoIbQTp7PjUGxZE5p0DUfaOByXct2D9AhQJBu9wE7q3aFcSCMoM8mE3jyACc40k_HkZHETOP64v6KJBWL7oo30WwteVNKux0bG-qBHnVGCvazUgibEEhyB4xdHmV98wjA095FRtslScoZg_tTe66UxuChkeCNhAFxgKdrkTXWrh4TLl-22uBN1xMGnmck4KijhOih5OVbu2EXjMYNuCQfJmqD7NzWqx4rBV679kOTFEGn2oAbQw&attredirects=0'));
  scrollpanel.add(grid);
  panel.add(scrollpanel);
  panel.add(button);
  app.add(panel);
  ss.show(app);
  return app; 
  }
}


function confirmSettings(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  howToPublishAsWebApp();
  return app;
}


function verifyPublished() {
  var scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl) {
    ScriptProperties.setProperty('scriptUrl', scriptUrl);
    return true;
  } else {
    return false;
  }
}


//UI to provide the URL to the gClassHub web application
// gClassHub is a web app that launches other scripts based on the student roster
function getGClassHubUrl() {
  var ssKey = ScriptProperties.getProperty('ssKey');
  var mode = ScriptProperties.getProperty('mode');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var LinkText = t("Custom link to gClassHub for this gClassFolders installation");
  var URL = 'https://script.google.com/macros/s/AKfycbxz9qFXlQrU_tPie5FCE27xWF1KbnSvYEpASezZ_OA2PXPs8ByV/exec?gClassKey=' + ssKey;
  var app = UiApp.createApplication();
  var panel = app.createVerticalPanel().setStyleAttribute('width','450px');
  var title = t("gClassHub: Take your gClassFolders experience to new heights with a gallery of add-on scripts, preconfigured and ready to use with your class roster!");
  var titleGrid = app.createGrid(1, 2);
  var icon = app.createImage(this.GCLASSLAUNCHERICONURL).setWidth("100px");
  var title = app.createLabel(title).setStyleAttribute('fontSize', '16px').setStyleAttribute('width','400px');
  titleGrid.setWidget(0, 0, icon);
  titleGrid.setWidget(0, 1, title);
  var alt = 'visit gClassHub for your classes';
  if (mode=="school") {
    alt = t("provide to teachers to visit gClassHub")
  }
  var labelText = t('Use the custom URL below to') + " " + alt;
  var prettyPanel = app.createDecoratedPopupPanel().setStyleAttribute('margin', '15px').setStyleAttribute('width','350px');
  var label = app.createLabel(labelText).setStyleAttribute('marginLeft', '15px').setStyleAttribute('marginTop', '10px');
  var anchor = app.createAnchor(LinkText, URL);
  panel.add(titleGrid);
  panel.add(label);
  prettyPanel.add(anchor);
  panel.add(prettyPanel);
  app.add(panel);
  ss.show(app);
  return app;
}


//Creates a NEW sheet, inserts required headings, and stores the 
//sheet Id for use elsewhere (eliminates trusting that "Active" sheet contains the roster)
//This only ever runs once in most instances.  Runs from menu on first use or if the user has deleted the roster sheet for some reason.
function createRosterSheet(properties){
  if (!properties) {
    var properties = ScriptProperties.getProperties();
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('gClassRoster');
  if ((!sheet)&&(!properties.sheetId)) {
    sheet = ss.insertSheet('gClassRoster');  
  }
  properties.sheetId = sheet.getSheetId();
  if (properties.mode == "school") {
    setgClassFoldersSchoolUid();
    setgClassFoldersSchoolSid();
  } else {  
    setgClassFoldersTeacherUid();
    setgClassFoldersTeacherSid();
  } 
  sheet.getRange("A1").setValue(t("Student First Name")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("B1").setValue(t("Student Last Name")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("C1").setValue(t("Student Email")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("D1").setValue(this.labels().class + t(" Name")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("E1").setValue(this.labels().period + " ~" + t("Optional") + "~").setComment(t("Don't change the name of this header!"));
  sheet.getRange("F1").setValue(t("Teacher Email(s)")).setComment(t("Don't change the name of this header!'"));
  SpreadsheetApp.flush();
  sheet.setFrozenRows(1);
  properties.sheetId = sheet.getSheetId();
  ScriptProperties.setProperties(properties);
  onOpen();  //refresh menu to change menu options
  ss.setActiveSheet(sheet);
}





//Function used to create folder Id Headings when the user runs the folder creation process
//If the user is re-running folder creation, this checks to see if the headings exist
function createFolderIdHeadings(){
  var sheet = getRosterSheet();
  var mode = ScriptProperties.getProperty('mode');
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf(t('Student ') + this.labels().dropbox + (' Folder Id'))==-1) {
    if (mode=="school") {
      sheet.getRange(1, lastCol+1, 1, 7).setValues([[t('Student') + ' ' + this.labels().dropBox + ' Id',t('Class Root Folder') + ' Id',t('Class View Folder') + ' Id',t('Class Edit Folder') + ' Id',t('Root Student Folder') + ' Id',t('Teacher Folder') + ' Id', t('Student') + ' ' + this.labels().class + ' ' + t('Folder') + ' Id']]).setComment(t("Don't manually change or delete any of these column headers or row values"));
    } else {
      sheet.getRange(1, lastCol+1, 1, 6).setValues([[t('Student ') + this.labels().dropBox + ' Id',t('Class Root Folder') + ' Id',t('Class View Folder') + ' Id',t('Class Edit Folder') + ' Id',t('Root Student Folder') + ' Id',t('Teacher Folder') + ' Id']]).setComment(t("Don't manually change or delete any of these column headers or row values"));
    }
    SpreadsheetApp.flush();
  }
}


//function prompts user to fix messed up headers in the sheet
function badHeaders() {
  var button = Browser.Buttons.YES_NO;
  if(Browser.msgBox(t("Required headers are are missing or impropertly labeled. Do you want the script to try fixing your headers?"), button))
  {
    fixHeaders();
  }
}




//function assigns headers to the sheet.  Headers are translated according to language and custom header settings.
function fixHeaders() {
  var sheet = getRosterSheet();
  sheet.getRange("A1").setValue(t("Student First Name")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("B1").setValue(t("Student Last Name")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("C1").setValue(t("Student Email")).setComment(t("Don't change the name of this header!"));
  sheet.getRange("D1").setValue(this.labels().class + t(" Name")).setComment(t("Class folders are created only for unique class names. Don\'t change the name of this header!"))
  sheet.getRange("E1").setValue(this.labels().period + " ~" + t("Optional") + "~").setComment(t("Don't change the name of this header!"));
  sheet.getRange("F1").setValue(t("Teacher Email(s)")).setComment(t("Don't change the name of this header!"));
  Browser.msgBox(t("gClassFolders has attempted to fix your headers.  Please check that everything in your roster sheet is as expected."));
}



function createClassFolders(){ //Create student folders
  //this step looks to see if the currently logged in user is looking to 
  //transfer ownership of folders to another teacher
  sortsheet();
  removeResumeTrigger();
  var lock = LockService.getPublicLock();
  lock.releaseLock();
  lock = LockService.getPublicLock();
  if (lock.tryLock(500)) {
    var startTime = new Date().getTime();
    var properties = ScriptProperties.getProperties();
    var dropBoxLabel = this.labels().dropBox;
    var dropBoxLabels = this.labels().dropBoxes;
    var periodLabel = this.labels().period;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssOwner = ss.getOwner().getEmail();
    var currUser = Session.getActiveUser().getEmail();
    var sheet = getRosterSheet(); 
    var dataRange = sheet.getDataRange().getValues();
    var labelObject = this.labels();
    var indices = returnIndices(dataRange, labelObject);
    var lang = properties.lang;
    var driveRoot = DocsList.getRootFolder();
    
    //get the top folder for active DBs, course folders, and archived course folders
    if (properties.mode == "school") {
      if (ssOwner!= currUser) {
        Browser.msgBox("In school mode, you can only run the folder creation process if are the owner of the spreadsheet and all student and teacher root folders. Please log in to the account used to set up gClassFolders")
        return;
      }
      //School mode will rely on top level folders to help organize the active teacher class roots, active student class roots, and archived course roots
      if ((!properties.topActiveClassFolderId)||(properties.topActiveClassFolderId=='')) {
        properties.topActiveClassFolderId = DocsList.createFolder(t('gClassFolders Active Teacher', lang) + " " + labelObject.classes).getId();
        ScriptProperties.setProperties(properties);
      }
      if ((!properties.topActiveDBFolderId)||(properties.topActiveDBFolderId=='')) {
        properties.topActiveDBFolderId = DocsList.createFolder(t('gClassFolders Active Student', lang) + " " + labelObject.class + " " + t('folders')).getId();
        ScriptProperties.setProperties(properties);
      }
      if ((!properties.topClassArchiveFolderId)||(properties.topClassArchiveFolderId=='')) {
        properties.topClassArchiveFolderId = DocsList.createFolder(t('gClassFolders Archived Teacher', lang) + " " + labelObject.classes).getId();
        ScriptProperties.setProperties(properties);
      }
      if ((!properties.topDBArchiveFolderId)||(properties.topDBArchiveFolderId=='')) {
        properties.topDBArchiveFolderId = DocsList.createFolder(t('gClassFolders Archived Student', lang) + " " + labelObject.dropBoxes).getId();
        ScriptProperties.setProperties(properties);
      }
      try {
        var topActiveClassFolder = DocsList.getFolderById(properties.topActiveClassFolderId);
        if (topActiveClassFolder.isTrashed()) {
          Browser.msgBox('The folder ' + topActiveClassFolder.getName() + ' is currently in your trash! Please return it to your drive before running the folder creation process.');
          return;
        }
        var topActiveDBFolder = DocsList.getFolderById(properties.topActiveDBFolderId);
        if (topActiveDBFolder.isTrashed()) {
          Browser.msgBox('The folder ' + topActiveDBFolder.getName() + ' is currently in your trash! Please return it to your drive before running the folder creation process.');
          return;
        }
        var topClassArchiveFolder = DocsList.getFolderById(properties.topClassArchiveFolderId);
        if (topClassArchiveFolder.isTrashed()) {
          Browser.msgBox('The folder ' + topClassArchiveFolder.getName() + ' is currently in your trash! Please return it to your drive before running the folder creation process.');
          return;
        }
        var topDBArchiveFolder = DocsList.getFolderById(properties.topDBArchiveFolderId);
        if (topDBArchiveFolder.isTrashed()) {
          Browser.msgBox('The folder ' + topDBArchiveFolder.getName() + ' is currently in your trash! Please return it to your drive before running the folder creation process.');
        }
      } catch(err) {
        Browser.msgBox("There was an error accessing your top level folders for active and archived courses. This can occur if for some reason you don't have access to one of these folders. If you need to regenerate these folders, you can delete their folder keys in the script properties. ");  
      }
      var studentRoots = getFolderRoots('sRoots');
    }
    
    //load the existing root folder ID info for teachers. These are stored in the scriptDb store for this script.
    //This means that deleting this script is a really bad idea if you want to keep the same root folders
    //for students and teachers over time
    var teacherRoots = getFolderRoots('tRoots');
    
    //This function adds robustness to the script by ensuring that we are looking in the correct
    //array indices for each of the elements.  If essential headers are missing, user is prompted to allow the script to auto-repair them.
    var indices = returnIndices(dataRange, labelObject);
    saveIndices(indices);
    writeProperties();
    //Sort by class, period, and last name to help consolidate rosters.  
    //note that this step is no longer technically necessary to ensure folder uniqueness.
    //left this in to provide ease of completeness check on class rosters
    sortsheet(indices.clsNameIndex,indices.clsPerIndex, indices.sLnameIndex);
    
    
    //now that all headings have been checked and indices identified
    //reload sheet and get 2D array of sheet data in case anything has changed.
    var sheet = getRosterSheet();
    var dataRange = sheet.getDataRange();
    dataRange = dataRange.getValues();
    
    //Remove wrap from non-header rows to economize on space.
    var wrapRange = sheet.getRange(1, 2, sheet.getLastRow(), sheet.getLastColumn()).setWrap(false);
    
    //Initialize counters
    var studentFoldersCreated = 0;
    
    var userEmail = Session.getEffectiveUser().getEmail(); //used later to check if script running user is the teacher whose email is listed, for ownership purposes 
    var clsFoldersCreated = []; //array to store all new folders created
    var editors = ss.getEditors();
    var editorEmails = [];
    for (var j=0; j<editors.length; j++) {
      editorEmails.push(editors[j].getEmail());
    }
    
    var interrupted = false;
    for (var i = 1; i < dataRange.length; i++) { //commence loop through all student class/period entries
      var loopStart = new Date().getTime();
      if ((loopStart - startTime)>310000) {
        setResumeTrigger(lock);
        interrupted = true;
        break;
      }
      var statusTagStudent = ""; //string used to concatenate student status messages
      var sFname = dataRange[i][indices.sFnameIndex]; // note that all sheet values are now 
      var sLname = dataRange[i][indices.sLnameIndex]; // addressed by variable index.  This is a safer
      var sEmail = dataRange[i][indices.sEmailIndex];  // way to roll than fixed indices, 
      var clsName = dataRange[i][indices.clsNameIndex]; // given how easy it is to drag a column in Google Spreadsheets
      var clsPer = dataRange[i][indices.clsPerIndex];    
      var tEmails = returnEmailAsArray(dataRange[i][indices.tEmailIndex]); //converts value of email column to an array of emails
      if (tEmails[0]=='') {
        tEmails[0]=userEmail;
        if (properties.mode = 'teacher') {
          sheet.getRange(i+1, indices.tEmailIndex+1).setValue(userEmail); //if email is blank, assume the person running the script is the teacher
        } else {
          Browser.msgBox(t("Row") + " " + i+1 + " " + t("is missing a teacher email address. Please fix this and restart the folder creation process from where it left off."));
        }
      }
      
      var sDropStatus = dataRange[i][indices.sDropStatusIndex];
      var tShareStatus = dataRange[i][indices.tShareStatusIndex];
      var rootStuFolderId = dataRange[i][indices.rsfIdIndex];
      var dropboxRootId = dataRange[i][indices.rsfIdIndex];
      var dropboxLabelId;
      if (i>0) {
        if ((dataRange[i][indices.clsNameIndex]==dataRange[i-1][indices.clsNameIndex])&&(dataRange[i][indices.clsPerIndex]!=dataRange[i-1][indices.clsPerIndex])&&(dataRange[i][indices.dbfIdIndex]=="")) {
          var dropbox = DocsList.getFolderById(dataRange[i-1][indices.rsfIdIndex]); //if student dropbox already exists in sheet 
          dropboxLabelId = dropbox.getParents()[0].getId();  //find its parent folder id in case a new period folder is needed
        }
      }
      var clsFolderId = dataRange[i][indices.crfIdIndex];
      var classViewId = dataRange[i][indices.cvfIdIndex];
      var classEditId = dataRange[i][indices.cefIdIndex];
      var teacherFId = dataRange[i][indices.tfIdIndex];
      
      if ((sDropStatus=="")&&(clsName!='')) { //only create folders in rows where students have blank status and a class assigned.
        var uniqueClasses = getUniqueClassNames(dataRange, indices.clsNameIndex, indices.crfIdIndex); //returns array of all classes that already have class root folders listed in the sheet
        if (uniqueClasses.indexOf(clsName)==-1) { //only create new class folder if this class folder doesn't already exist
          try {
            var clsFolder = DocsList.createFolder(clsName);
          } catch(err)  {  
            if (err.message.search("too many times")>0) {
              Browser.msgBox(t("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard", lang));
              return;
            }
          }
          try {
            gClassFolders_logTeacherClassFolderCreated();
          } catch(err) {
          }
          var clsFolderId = clsFolder.getId();
          dataRange[i][indices.crfIdIndex] = clsFolderId;
          clsFoldersCreated.push(clsName);
          var classEdit = clsName +" - " + t("Edit", lang);  
          var classView = clsName +" - " + t("View", lang); 
          var teacherFolderLabel = clsName + " - " + t("Teacher", lang);
          var tMessage = t("Folders created for", lang) + " " + clsName;
          //treat the first listed teacher email as primary...allow secondary teachers to be added
          for (var j=0; j<tEmails.length; j++) {//Transfer ownership of rootFolder to teacher if teacher email is designated.  Check that designated email is not the user running the script.
            try {
              if (properties.mode=='school') {
                //teacherRoots is an array of objects, retrieved from scriptDb store, containing the folder Ids of all teacher root folders (active and archived).
                teacherRoots = moveToTeacherRoot(teacherRoots, tEmails[j], clsFolderId, topActiveClassFolder, topClassArchiveFolder, 'active', lang, driveRoot);
                //function above is used to create new teacher roots for users that don't yet have them and transfer a given folder
                //into the active or archive root folder.  This is only ever used in school mode.
              } 
              if ((tEmails[j] != "")&&(tEmails[j] != userEmail)){
                DocsList.getFolderById(clsFolderId).addEditor(tEmails[j]); 
                tMessage += ", " + tEmails[j] + " " + t("added as editor.");
              } else { //do this if teacher email is the same as that of the script user, or if tEmail is blank. This can only happen in teacher mode.
                tMessage += t(", you're the teacher.", lang);
                sheet.getRange(i+1, indices.tShareStatusIndex+1).setValue(tMessage);  
              }
            } catch(err) {
              DocsList.getFolderById(clsFolderId).addEditor(tEmails[j]);
              tMessage += t(", Error sharing folder for: ", lang) + tEmails[j] + t("Error: ") + err;
            }
            sheet.getRange(i+1, indices.tShareStatusIndex+1).setValue(tMessage);
            if ((tEmails[j]!=userEmail)&&(editors.indexOf(tEmails[j])==-1)) {
              ss.addViewer(tEmails[j]);
            }
          }
          //Create class edit, class view, and dropbox sub-folders
          try {
            var classEditId = DocsList.getFolderById(clsFolderId).createFolder(classEdit).getId();
            var classViewId = DocsList.getFolderById(clsFolderId).createFolder(classView).getId();
            var teacherFId = DocsList.getFolderById(clsFolderId).createFolder(teacherFolderLabel).getId();
          } catch(err) {  
            if (err.message.search("too many times")>0) {
              Browser.msgBox(t("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard"));
              return;
            }
          }
          dataRange[i][indices.tfIdIndex] = teacherFId;
          rootStuFolderId = DocsList.getFolderById(clsFolderId).createFolder(dropBoxLabels).getId(); //assign rootStuFolderId for now, pending a check whether period exists
          for (var j=0; j<tEmails.length; j++) {
            if ((tEmails[j]!="")&&(tEmails[j] != userEmail)) {//execute only if teacher email field is neither blank nor the same as the user running the script
              try {
                DocsList.getFolderById(classEditId).addEditor(tEmails[j]);
                DocsList.getFolderById(classViewId).addEditor(tEmails[j]);
                DocsList.getFolderById(rootStuFolderId).addEditor(tEmails[j]);
                DocsList.getFolderById(teacherFId).addEditor(tEmails[j]);
              } catch (err) {
                tMessage += t(", Error sharing folder for: ", lang) + tEmails[j] + t("Error: ", lang) + err;
              }        
            }
          }
          var dropboxLabelId = rootStuFolderId;
          //move to next username in class
        } // End of create class Folders
        var classRoster = null;
        var perRoster = null;
        if(rootStuFolderId=="") {
          perRoster = getClassRoster(dataRange, indices, clsName, clsPer);
          rootStuFolderId =  getClassFolderId(perRoster, indices.rsfIdIndex);
        }  
        if ((!dropboxLabelId)||(dropboxLabelId=="")) {
          classRoster = getClassRoster(dataRange, indices, clsName);
          if (!rootStuFolderId) {  
            rootStuFolderId =  getClassFolderId(classRoster, indices.rsfIdIndex);
          }      
          var dropboxRoot = DocsList.getFolderById(rootStuFolderId);
          dropboxLabelId = dropboxRoot.getId();
        }
        if (clsFolderId=="") {
          if (!classRoster) {
            classRoster = getClassRoster(dataRange, indices, clsName);
          }
          clsFolderId =  getClassFolderId(classRoster, indices.crfIdIndex);
        }
        if (classViewId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          classViewId = getClassFolderId(classRoster, indices.cvfIdIndex);
        }
        if (classEditId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          classEditId = getClassFolderId(classRoster, indices.cefIdIndex);
        }
        if (teacherFId=="") {
          classRoster = getClassRoster(dataRange, indices, clsName);
          teacherFId = getClassFolderId(classRoster, indices.tfIdIndex);
        }
        //If a class period is chosen, look to see if it is new or already existing
        if (clsPer != "") {
          var uniqueClasses = getUniqueClassPeriods(dataRange, indices.clsNameIndex, indices.clsPerIndex, indices.rsfIdIndex, labelObject); //get unique ClassPer as array
          if (uniqueClasses.indexOf(clsName + " " + periodLabel + " " + clsPer)==-1) { //look to see if this row's ClassPer exists in the array.  If not make a new student dropbox folder for the period
            rootStuFolderId = DocsList.getFolderById(dropboxLabelId).createFolder(clsName + " " + periodLabel + " " + clsPer + " " + dropBoxLabels).getId();
            clsFoldersCreated.push(clsName + " " + periodLabel + " " + clsPer);
            for (var j=0; j<tEmails.length; j++) {
              if ((tEmails[j] != "")&&(tEmails[j] != userEmail)) {
                DocsList.getFolderById(rootStuFolderId).addEditor(tEmails[j]);
              }
            }
          }
        } // End if Per
        
        //Create students
        var dbfId = dataRange[i][indices.dbfIdIndex];
        var studentFolderObj = createDropbox(sLname,sFname,sEmail,clsName,classEditId,classViewId,rootStuFolderId,tEmails, userEmail, properties, dropBoxLabel, lang);
        if (properties.mode == 'school') {
          studentRoots = moveToStudentRoot(studentRoots, sEmail, sFname, sLname, studentFolderObj.studentClassRoot, topActiveDBFolder, topDBArchiveFolder, 'active', lang, driveRoot);
        }
        studentFoldersCreated++;
        var values = [];
        values[0] = [];
        dataRange[i][indices.dbfIdIndex] = studentFolderObj.studentDropboxId;
        dataRange[i][indices.crfIdIndex] = clsFolderId;
        dataRange[i][indices.cvfIdIndex] = classViewId; 
        dataRange[i][indices.cefIdIndex] = classEditId;
        dataRange[i][indices.rsfIdIndex] = rootStuFolderId;
        dataRange[i][indices.tfIdIndex] = teacherFId;
        if (properties.mode == 'school') {
          dataRange[i][indices.scfIdIndex] = studentFolderObj.studentClassRootId;
        }
        
        values[0].push('=hyperlink("'+ studentFolderObj.studentDropbox.getUrl() +'";"'+studentFolderObj.studentDropboxId + '")');
        values[0].push('=hyperlink("'+ DocsList.getFolderById(clsFolderId).getUrl() + '";"' + clsFolderId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.classView.getUrl() + '";"' + classViewId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.classEdit.getUrl() + '";"' + classEditId + '")');
        values[0].push('=hyperlink("' + studentFolderObj.rootStudentFolder.getUrl() + '";"' + rootStuFolderId + '")');
        values[0].push('=hyperlink("' + DocsList.getFolderById(teacherFId).getUrl() + '";"' + teacherFId + '")');
        if (properties.mode == 'teacher') {
          sheet.getRange(i+1, indices.dbfIdIndex + 1, 1, 6).setValues(values).setFontColor('black');
        } else {
          values[0].push('=hyperlink("'+ studentFolderObj.studentClassRoot.getUrl() +'";"'+studentFolderObj.studentClassRootId +'")');
          sheet.getRange(i+1, indices.dbfIdIndex + 1, 1, 7).setValues(values).setFontColor('black');
        }
        
        //add Status 
        sheet.getRange(i+1, indices.sDropStatusIndex+1).setValue(studentFolderObj.statusTagF).setFontColor('black');
        SpreadsheetApp.flush();
      }
    }//end loop through all student class/period entries
    
    var msg = '';
    if (interrupted) {
      msg += t('The folder creation process was interrupted and will restart automatically to avoid script timeout. Please allow the script at least 1 minute to resume before attempting to resume manually. ', lang);
    }
    if (clsFoldersCreated.length>0) {
      msg = t("Class folders were created for:", lang) + "\\n" + clsFoldersCreated.join(", \\n") + "\\n \\n";
    } 
    if (studentFoldersCreated>0) {
      ScriptProperties.setProperty('alreadyRan', 'true');
      onOpen();
      msg += " " + studentFoldersCreated + t(" new ", lang) + dropBoxLabels + t(" were created.", lang) + "\\n \\n";
    } else {
      msg += t("No new folders were created.  Folders are only created for rows with a blank", lang) + "\"" + t("Status: Student Dropbox", lang) + "\" " + t("value", lang) + "\\n \\n";
    }
    lock.releaseLock();
    Browser.msgBox(msg);
  } else {
    Browser.msgBox(t("It appears the folder creation process is already underway. Please don't interrupt!", lang));
  }
}






function createDropbox(sLnameF,sFnameF,sEmailF,clsNameF,classEditIdF,classViewIdF,rootStuFolderId,tEmails,userEmail, properties, dropboxLabel, lang) {
  var returnObject = new Object();
  if (properties.mode == 'school') {
    try {
      var studentClassRoot = DocsList.createFolder(clsNameF);
      returnObject.studentClassRootId = studentClassRoot.getId();
      studentClassRoot.addViewer(sEmailF);
    } catch(err) {
      if (err.message.search("too many times")>0) {
        Browser.msgBox(t("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard", lang));
        return;
      }
    }
  }
  var dropboxNameF = sLnameF + ", " + sFnameF + " - " + clsNameF + " - " + dropboxLabel;
  var rootStudentFolder = DocsList.getFolderById(rootStuFolderId);
  var studentDropbox = rootStudentFolder.createFolder(dropboxNameF);
  returnObject.statusTagF = dropboxLabel + t(" created", lang);
  try {
    var classEdit = DocsList.getFolderById(classEditIdF);
    var classView = DocsList.getFolderById(classViewIdF);
    if (properties.mode == 'school') {
      studentDropbox.addToFolder(studentClassRoot);
      classEdit.addToFolder(studentClassRoot);
      classView.addToFolder(studentClassRoot);
    }
    classEdit.addEditor(sEmailF);
    classView.addViewer(sEmailF);
    studentDropbox.addEditor(sEmailF);
    returnObject.statusTagF += t(", and shared with ", lang) + sEmailF;
  } catch(e) {
    Logger.log(t("Error with email", lang) + " (" + sEmailF + "). " + e);
    returnObject.statusTagF += t(", Error with Student email: folder created but not shared", lang); 
  }
  var studentDropboxId = studentDropbox.getId()
  returnObject.studentDropboxId = studentDropboxId;
  returnObject.studentDropbox = studentDropbox;
  returnObject.classEdit = classEdit;
  returnObject.classView = classView;
  returnObject.studentClassRoot = studentClassRoot;
  returnObject.rootStudentFolder = rootStudentFolder;
  for (var j=0; j<tEmails.length;j++) {
    if ((tEmails[j] != "")&&(tEmails[j] != userEmail)) {  
      try {
        DocsList.getFolderById(studentDropboxId).addEditor(tEmails[j]); 
        returnObject.statusTagF += t(", editing rights added for ", lang) + tEmails[j];
      } catch(err) {
        returnObject.statusTagF += t(", error giving editing rights to ", lang) + tEmails[j] + "." + err;
      }
    }
  } 
  return returnObject;
}
