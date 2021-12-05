const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const sourceDirPath = path.join(__dirname, '../../temp_file');

/**
 * Converting sheet data in JSON format
 * @param {} sheetName 
 * @param {*} headerIndex 
 * @returns 
 */
const convertSheetToJson = (sheetName, headerIndex) => {
  try {
    let config = {
      header: "A"
    }
    let sheetData = xlsx.utils.sheet_to_json(sheetName, config);
    if (sheetData && sheetData.length > 0) {
      const header = sheetData[headerIndex];
      let loopIndex = headerIndex + 1;
      let jsonRows = Array();
      for (loopIndex; loopIndex < sheetData.length; loopIndex++) {
        let tempObj = {};
        for (const property in sheetData[loopIndex]) {
          tempObj[header[property]] = sheetData[loopIndex][property];
        }
        jsonRows.push(tempObj);
      }
      return jsonRows;
    }
  } catch (error) {
    throw error;
  }
}

/**
 * Processing the data of Address sheet
 * @param {*} addressSheet 
 */
 const processAddressSheet = (addressSheet) => {
  try {
    const headerStartIndexInAddressSheet = 3;//based on position of header in the sheet
    let addressJsonData = convertSheetToJson(addressSheet, headerStartIndexInAddressSheet);
    for(let data of addressJsonData) {
      console.log(data);
    }
  } catch (error) {
    throw error;
  }
}

/**
 * Processing the data of Candidate sheet
 * @param {*} candidateSheet 
 */
const processCandidateSheet = (candidateSheet) => {
  try {
    const headerStartIndexInCandidateSheet = 2;//based on position of header in the sheet
    let candidateJsonData = convertSheetToJson(candidateSheet, headerStartIndexInCandidateSheet);
    for(let data of candidateJsonData) {
      console.log(data);
    }
  } catch (error) {
    throw error;
  }
}

/**
 * Creating the list of names of files present directly in temp_file dir
 * @returns 
 */
const getFiles = () => {
  try {
    let files = Array();
    fs.readdirSync(sourceDirPath).forEach(file => {
      if (!fs.lstatSync(path.resolve(sourceDirPath, file)).isDirectory()) {
        files.push(file);
      }
    });
    return files;
  } catch (error) {
    throw error;
  }  
}

/**
 * Getting the list of name of files present directly in temp_file folder
 * Iterating this list of name of files and then processing each sheet in each file
 * Each sheet is being processed by corresponding function
 * @returns 
 */
const readTestFile = async () => {
  try {
    const files = getFiles();
    if(files && files.length === 0) {
      console.log("No file present");
      return true;
    }
    const filePath = sourceDirPath + "/" + files[0];
    //Reading the excel file candidate.xlsx
    const file = xlsx.readFile(filePath);
    if(file && file.SheetNames && file.SheetNames.length > 0) {
      for(let sheetName of file.SheetNames) {
        if(sheetName === 'candidates') {
          processCandidateSheet(file.Sheets[sheetName]);
        } else if(sheetName === 'address') {
          processAddressSheet(file.Sheets[sheetName]);
        }
      }
    }
  } catch (error) {
    throw error;
  }  
}

const initiate = async () => {
  await readTestFile();
}

setTimeout(() => {
  initiate();
}, 2000);