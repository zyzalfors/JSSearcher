let FSObj = new ActiveXObject("Scripting.FileSystemObject");
let textArea = document.querySelector("textarea");

function formatPath(path) {
 return (path.trim() + "\\").replace(/\\+/g,"\\");
}

function enableAdvancedOptions() {
 if(document.getElementById("enableAdvanced").checked) {
  document.getElementById("regex").disabled = false;
  document.getElementById("caseIns").disabled = false;
  document.getElementById("minSize").disabled = false;
  document.getElementById("maxSize").disabled = false;
  document.getElementById("searchFiles").disabled = false;
  document.getElementById("searchDir").disabled = false;
  document.getElementById("scanSub").disabled = false;
  document.getElementById("moreInfo").disabled = false;
  document.getElementById("search").disabled = false;
  document.getElementById("quick").disabled = true;
 }
 else {
  document.getElementById("regex").disabled = true;
  document.getElementById("caseIns").disabled = true;
  document.getElementById("minSize").disabled = true;
  document.getElementById("maxSize").disabled = true;
  document.getElementById("searchFiles").disabled = true;
  document.getElementById("searchDir").disabled = true;
  document.getElementById("scanSub").disabled = true;
  document.getElementById("moreInfo").disabled = true;
  document.getElementById("search").disabled = true;
  document.getElementById("quick").disabled = false;
 }
}

function quickBrowse() {
 let targetDirPath = formatPath(document.getElementById("targetDirPath").value);
 if(FSObj.FolderExists(targetDirPath) && targetDirPath !== "\\") {
  let report = "Files:\n";
  let files = new Enumerator(FSObj.GetFolder(targetDirPath).Files);
  for(;!files.atEnd() ;files.moveNext()) {
   let file = files.item();
   report += file.Name + " , " + file.Size + " bytes\n";
  }
  report += "Directories:\n";
  let directories = new Enumerator(FSObj.GetFolder(targetDirPath).SubFolders);
  for(;!directories.atEnd() ;directories.moveNext()) {
   let directory = directories.item();
   let dirSize = 0;
   try {
    dirSize = parseFloat(directory.Size); 
   }
   catch(error) {  }
   report += directory.Name + " , " + dirSize + " bytes\n"; 
  }
  textArea.style.display = "inline";
  textArea.value = report.slice(0,-1);
  document.getElementById("reportDirPath").disabled = false;
  document.getElementById("save").disabled = false;
  document.getElementById("filePath").disabled = false;
  document.getElementById("open").disabled = false;
 }
 else {
  alert("Target directory not found"); 
 }
} 
 
function openFile() {
 let SObj = new ActiveXObject("Wscript.Shell");
 let filePath = formatPath(document.getElementById("filePath").value).slice(0, -1);
 if(FSObj.FileExists(filePath)) {
  try { 
   SObj.run("\"" + filePath + "\""); 
  }
  catch(error) { 
   window.open(filePath);
  }
 }
 else { 
  alert("File not found"); 
 }
}

function getRegex(regex, caseIns) {
 try { 
  return caseIns ? new RegExp(regex, "i") : new RegExp(regex); 
 }
 catch(error) {
  alert("Invalid regular expression");
  return new RegExp(".+");
 }
}

function saveReport() {
 let reportDirPath = formatPath(document.getElementById("reportDirPath").value);
 if(FSObj.FolderExists(reportDirPath) && reportDirPath !== "\\") {
  let reportFilePath = reportDirPath + "JSSearcherReport.txt";
  try {
   let text = textArea.value;
   let reportFile = FSObj.CreateTextFile(reportFilePath, true, true);
   reportFile.Write(text);
   reportFile.Close();
   alert("Report file saved");
  }
  catch(error) { 
   alert("Report file cannot be saved"); 
  }
 }
 else { 
  alert("Report directory not found"); 
 }
}

function addToReport(name, size, regExp, minSize, maxSize) {
 let matches = name.match(regExp);
 minSize = isNaN(minSize) ? size : minSize;
 maxSize = isNaN(maxSize) ? size : maxSize;
 return (matches !== null ? matches.indexOf(name) !== -1 : false) && size >= minSize && size <= maxSize;
}

function searchAndWrite(targetDirPath, regExp, minSize, maxSize, searchFiles, searchDir, scanSub, moreInfo, reportObj) {
 let files = new Enumerator(FSObj.GetFolder(targetDirPath).Files);
 let directories = new Enumerator(FSObj.GetFolder(targetDirPath).SubFolders);
 if(searchFiles) {
  for(;!files.atEnd();files.moveNext()) {
   reportObj.scanned++;
   let file = files.item();
   let fileName = file.Name;
   let fileSize = parseFloat(file.Size);
   if(addToReport(fileName, fileSize, regExp, minSize, maxSize)) {
    reportObj.nFiles++;
    if(moreInfo) { 
     reportObj.report += reportObj.nFiles + " | " + file.Type + " | " + targetDirPath + " | " + fileName + " | " + fileSize + " | " + new Date(file.DateLastModified).toLocaleString() + " | " + new Date(file.DateCreated).toLocaleString() + " | " + new Date(file.DateLastAccessed).toLocaleString() + "\n";
    }
    else { 
     reportObj.report += reportObj.nFiles + " | File | " + targetDirPath + " | " + fileName + "\n"; 
    }
   }
  }
 }
 for(;!directories.atEnd();directories.moveNext()) {
  let directory = directories.item();
  let dirName = directory.Name;
  if(searchDir) {
   reportObj.scanned++;
   let dirSize = 0;
   try {
    dirSize = parseFloat(directory.Size); 
   }
   catch(error) {  }
   if(addToReport(dirName, dirSize, regExp, minSize, maxSize)) {
    reportObj.nDirs++;
    if(moreInfo) { 
     reportObj.report += reportObj.nDirs + " | " + directory.Type + " | " + targetDirPath + " | " + dirName + " | " + dirSize + " | " + new Date(directory.DateLastModified).toLocaleString() + " | " + new Date(directory.DateCreated).toLocaleString() + " | " + new Date(directory.DateLastAccessed).toLocaleString() + "\n"; 
    }
    else { 
     reportObj.report += reportObj.nDirs + " | Directory | " + targetDirPath + " | " + dirName + "\n"; 
    }
   }
  }
  if(scanSub) { 
   searchAndWrite(formatPath(targetDirPath + dirName), regExp, minSize, maxSize, searchFiles, searchDir, scanSub, moreInfo, reportObj); 
  }
 }
}

function searchElements() {
 let regex = document.getElementById("regex").value;
 let minSize = parseFloat(document.getElementById("minSize").value);
 let maxSize = parseFloat(document.getElementById("maxSize").value);
 let caseIns = document.getElementById("caseIns").checked;
 let searchFiles = document.getElementById("searchFiles").checked;
 let searchDir = document.getElementById("searchDir").checked;
 let scanSub = document.getElementById("scanSub").checked;
 let moreInfo = document.getElementById("moreInfo").checked;
 let regExp = getRegex(regex, caseIns);
 let targetDirPath = formatPath(document.getElementById("targetDirPath").value);
 if(FSObj.FolderExists(targetDirPath) && targetDirPath !== "\\") {
  let reportObj = {nFiles: 0, nDirs: 0, scanned: 0, report: ""};
  let start = new Date();
  searchAndWrite(targetDirPath, regExp, minSize, maxSize, searchFiles, searchDir, scanSub, moreInfo, reportObj);
  let elapsed = (new Date() - start)/1000;
  let report = "JS Searcher\n"
               + "Version: 1.1.7\n"
	       + "Local date and time: " + new Date().toLocaleString() + "\n"
               + "Time elapsed: " + elapsed + " s\n" 
	       + "Target directory: " + targetDirPath + "\n" 
	       + "Regular expression: " + regex + "\n"
	       + "Minimum size: " + (isNaN(minSize) ? "undefined" : minSize + " byte") + "\n"
	       + "Maximum size: " + (isNaN(maxSize) ? "undefined" : maxSize + " byte") + "\n"
	       + "Case insensitive: " + caseIns + "\n"
	       + "Search files: " + searchFiles + "\n" 
	       + "Search directories: " + searchDir + "\n" 
	       + "Scan subdirectories: " + scanSub + "\n"
	       + "Detailed info: " + moreInfo + "\n"
	       + "Found files: " + reportObj.nFiles + "\n"  
	       + "Found directories: " + reportObj.nDirs + "\n"
	       + "Total: " + (reportObj.nFiles + reportObj.nDirs) + "\n"
	       + "Scanned items: "+ reportObj.scanned;
  if(reportObj.nFiles !== 0 || reportObj.nDirs !== 0) {
   report += "\n\n  | Type | Parent directory | Name";
   report += moreInfo ? " | Size (byte) | Date last modified | Date created | Date last accessed\n" : "\n";
   report += reportObj.report.slice(0,-1);
  }
  else {
   report += "\n\nNo results found"; 
  }
  textArea.style.display = "inline";
  textArea.value = report;
  document.getElementById("reportDirPath").disabled = false;
  document.getElementById("save").disabled = false;
  document.getElementById("filePath").disabled = false;
  document.getElementById("open").disabled = false;
 }
 else {
  alert("Target directory not found"); 
 }
}
