//type must be "tRoots" or "sRoots"
function getFolderRoots(type) {
  var db = ScriptDb.getMyDb();
  var result = db.query({type: type});
  var results = []
  while (result.hasNext()) {
    var current = result.next();
    results.push(current);
  }
  return results;
}

function doGet(e) {
  var app = UiApp.createApplication();
  var activeFolderId = e.parameter.activeFolderId;
  var archiveFolderId = e.parameter.archiveFolderId;
  var root = DocsList.getRootFolder();
  var message = '';
  try {
    DocsList.getFolderById(activeFolderId).addToFolder(root);
    DocsList.getFolderById(archiveFolderId).addToFolder(root);
    message = t('Success! Your active and archived course folders have been moved to your Drive.');
  } catch (err) {
    message = t(err.message);
    message += t(" Please contact your gClassFolders admin to let them know something is wrong;(");
  }
  app.add(app.createLabel(message).setStyleAttribute('margin', '25px'));
  return app;  
}



function makeStudentRoots(sEmail, sFName, sLName, topActiveStudentFolder, topArchiveStudentFolder) {
  try {
    var activeFolder = topActiveStudentFolder.createFolder(sLName +", "+ sFName + t(" Current Classes"));
    var archiveFolder = topArchiveStudentFolder.createFolder(sLName +", "+ sFName + t(" Class Work Archives"));
  } catch(err) {  
    if (err.message.search("too many times")>0) {
      Browser.msgBox(t("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard"));
      return;
    }
  }
  var activeFolderId = activeFolder.getId();
  var archiveFolderId = archiveFolder.getId();
  activeFolder.addViewer(sEmail);
  archiveFolder.addViewer(sEmail);
  var db = ScriptDb.getMyDb();
  var studentRoots = {type: 'sRoots', email: sEmail, activeFolderId: activeFolderId, archiveFolderId: archiveFolderId};
  db.save(studentRoots);
  var scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl) {
    var body = t("Course folders have recently been created and shared with you by ") + this.userEmail + "."; 
    body += t("One folder is for your 'Active' classes, the other is for ");
    body += t("'Archived,' or old classes, and will be used in the future to keep your old work organized.") + "<br>";
    body += t("Please") + "<a href=\"" + scriptUrl + "?activeFolderId=" + activeFolderId + "&archiveFolderId=" + archiveFolderId + "\">" + t('click this link') + "</a> " + t("and you authorize the script to run, and you should see these folders added to your Drive.");
    body += "<br>" + t("Note: You may need to refresh your browser once for this to work.");
    MailApp.sendEmail(sEmail, t('Action required: Please add your class folders to your Drive'),'', {htmlBody: body})
  }
  return studentRoots;
}


function makeTeacherRoots(tEmail, topActiveTeacherFolder, topArchiveTeacherFolder) {
  try {
    var activeFolder = topActiveTeacherFolder.createFolder(tEmail + t(" Active Classes"));
    var archiveFolder = topArchiveTeacherFolder.createFolder(tEmail + t(" Class Archives"));
  } catch(err) {  
    if (err.message.search("too many times")>0) {
      Browser.msgBox(t("You have exceeded your account quota for creating Folders.  Try waiting 24 hours and continue running from where you left off. For best results with this script, be sure you are using a Google Apps for EDU account. For quota information, visit https://docs.google.com/macros/dashboard"));
      return;
    }
  }
  var activeFolderId = activeFolder.getId();
  var archiveFolderId = archiveFolder.getId();
  activeFolder.addViewer(tEmail);
  archiveFolder.addViewer(tEmail);
  var db = ScriptDb.getMyDb();
  var teacherRoots = {type: 'tRoots', email: tEmail, activeFolderId: activeFolderId, archiveFolderId: archiveFolderId};
  db.save(teacherRoots);
  var scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl) {
    var body = t("Course folders have recently been created and shared with you by ") + this.userEmail + "."; 
    body += t("One folder is for your 'Active' classes, the other is for ");
    body += t("'Archived,' or old classes, and will be used in the future to keep your old work organized.") + "<br>";
      body += t("Please") + "<a href=\"" + scriptUrl + "?activeFolderId=" + activeFolderId + "&archiveFolderId=" + archiveFolderId + "\">" + t('click this link') + "</a> " + t("and you authorize the script to run, and you should see these folders added to your Drive.");
      body += "<br>" + t("Note: You may need to refresh your browser once for this to work.");
      MailApp.sendEmail(tEmail, t('Action required: Please add your class folders to your Drive'),'', {htmlBody: body})
    }
  return teacherRoots;
}


// Here type is 'active' or 'archive'
// folderId is the if of the folder to be added to the student root folder 
// topFolder is the folder on the script owner account to organize roots under
// results is the array of db objects containing existing student roots
function moveToStudentRoot(results, sEmail, sFName, sLName, folderId, topActiveDBFolder, topDBArchiveFolder, type) {
  var found = false;
  var activeRootId = '';
  var archiveRootId = '';
  var folder = DocsList.getFolderById(folderId);
  for (var i=0; i<results.length; i++) {
    if (results[i].email == sEmail) {
      found = true;
      var studentRootFolders = results[i];
      break;
    }
  }
  if (!found) {
    var studentRootFolders = makeStudentRoots(sEmail, sFName, sLName, topActiveDBFolder, topDBArchiveFolder);
    results.push(studentRootFolders);
  }
  //need to add method for dealing with trashed student roots
  var activeFolder = DocsList.getFolderById(studentRootFolders.activeFolderId);
  if (type == 'archive') {
    var archiveFolder = DocsList.getFolderById(studentRootFolders.archiveFolderId);
    folder.addToFolder(archiveFolder); 
    folder.removeFromFolder(activeFolder);
    folder.removeEditor(sEmail);
    folder.addViewer(sEmail);
    try {
      gClassFolders_logStudentClassFolderArchived();
    } catch(err) {
    }
  }
  if (type == 'active') {
    var activeFolder = DocsList.getFolderById(studentRootFolders.activeFolderId);
    folder.addToFolder(activeFolder);
    folder.removeFromFolder(DocsList.getRootFolder());
  }
  return results;
}


// Here type is 'active' or 'archive'
// folder is the folder to be added to the student root folder 
// topFolder is the folder on the script owner account to organize roots under
// results is the array of db objects containing existing student roots
function moveToTeacherRoot(results, tEmail, folderId, topActiveTeacherFolder, topArchiveTeacherFolder, type) {
  var found = false;
  var folder = DocsList.getFolderById(folderId);
  var activeRootId = '';
  var archiveRootId = '';
  for (var i=0; i<results.length; i++) {
    if (results[i].email == tEmail) {
      found = true;
      var teacherRootFolders = results[i];
      break;
    }
  }
  if (!found) {
    var teacherRootFolders = makeTeacherRoots(tEmail, topActiveTeacherFolder, topArchiveTeacherFolder);
    results.push(teacherRootFolders);
  }
  //need to add method for dealing with trashed teacher roots
  var activeFolder = DocsList.getFolderById(teacherRootFolders.activeFolderId);
  
  if (type == 'archive') {
    var archiveFolder = DocsList.getFolderById(teacherRootFolders.archiveFolderId);
    folder.removeFromFolder(activeFolder);
    folder.removeEditor(tEmail);
    folder.addViewer(tEmail);
    folder.addToFolder(archiveFolder); 
  }
  if (type == 'active') { 
    folder.addToFolder(activeFolder);
    folder.removeFromFolder(DocsList.getRootFolder());
  }
  return results;
}



function archiveToTeacherRoot(results, sEmail, sFName, sLName, folderToArchive) {
  var found = false;
  var archiveRootId = '';
  for (var i=0; i<results.length; i++) {
    if (results[i].email = sEmail) {
      found = true;
      archiveRootId = results[i].archiveRootId;
      break;
    }
  }
  if (!found) {
    archiveRootId = makeStudentArchiveRoot(sEmail, sFName, sLName, topFolder);
  }
  var archiveFolder = DocsList.getFolderById(archiveRootId);
  folderToArchive.addToFolder(archiveFolder); 
}



//Never run this unless you absolutely intend to wipe out all folder associations for student and teacher root (Active and Archived) folders
function deleteAll() {
  var db = ScriptDb.getMyDb();
  while (true) {
    var result = db.query({}); // get everything, up to limit
    if (result.getSize() == 0) {
      break;
    }
    while (result.hasNext()) {
      db.remove(result.next());
    }
  }
}


function resendRootEmails() {
  var roots = getFolderRoots('sRoots');
  debugger;
  for (var i=0; i<roots.length; i++) {
    var email =  roots[i].email;
    var activeFolderId = roots[i].activeFolderId;
    var archiveFolderId = roots[i].archiveFolderId;
    var scriptUrl = ScriptApp.getService().getUrl();
    if (scriptUrl) {
      var body = t("OK: Let's try that again.  (Andrew here, beta testing something for this afternoon) Course folders have recently been created and shared with you by ") + this.userEmail + "."; 
      body += t("One folder is for your 'Active' classes, the other is for ");
      body += t("'Archived,' or old classes, and will be used in the future to keep your old work organized.") + "<br>";
      body += t("Please") + " " + "<a href=\"" + scriptUrl + "?activeFolderId=" + activeFolderId + "&archiveFolderId=" + archiveFolderId + "\">" + t('click this link') + "</a> " + " " + t("and you authorize the script to run, and you should see these folders added to your Drive.");
      body += "<br>" + t("Note: You may need to refresh your browser once for this to work.");
      MailApp.sendEmail(email, t('Action required (2nd attempt): Please add your class folders to your Drive'),'', {htmlBody: body})
    }
  }
}
