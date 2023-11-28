// **********************************************************
// *** EDITABLE *** Please change any of the following values
// **********************************************************

//name of the worksheet to find the tickets to push to Dynatrace
let dataWorksheetName = "Sheet1";

//the start count of the records, 1 skips the column headers in the worksheet
let itemStartCount = 1;

//max total number of records to send to Dynatrace API, 0 if you want to include all records in the worksheet
let maxItemCount = 0;

//max number of records to send to Dynatrace API at a time (bundles)
let apiItemCount = 1000;

//Dynatrace live id 1st part of the Dynatrace live url
//Example 1st part is <liveid> in the following format https://<liveid>.live.dynatrace.com/
let dtId = "";

//Dynatrace API token that has the ingest logs scope
let dtToken = "";

//Show or not show the console messages
let showConsoleMessages = true;

// *******************************************************
// *** READONLY *** do not change anything after this note
// *******************************************************

//active workbook
let activeWorkbook: ExcelScript.Workbook = null;

//name of the worksheet to find the tickets to push to Dynatrace
let logWorksheetName = "Logs";

const rows: (string | boolean | number)[][] = [];
let apiUrl = 'https://' + dtId + '.live.dynatrace.com/api/v2/logs/ingest';
let apiToken = 'Api-Token ' + dtToken;
let startTime: Date = new Date();

const byteLengthUTF16 = (str: string) => str.length * 2
const byteLengthUTF8 = (str: string) => new Blob([str]).size
const units = ['bytes', 'KiB', 'MiB', 'GiB', 'TiB', 'PiB', 'EiB', 'ZiB', 'YiB'];


async function main(workbook: ExcelScript.Workbook): Promise<void> {

    activeWorkbook = workbook;
    await pushData();
    await saveLogRecords();

    return null;
}

async function pushData(): Promise<object[]> {
  
  let returnObjects: TableData[] = [];
  const dataWorksheet = activeWorkbook.getWorksheet(dataWorksheetName);

  if (dataWorksheet) {
    if (showConsoleMessages) console.log('Found the ' + dataWorksheetName + ' worksheet');
    const dataTables = dataWorksheet.getTables();
    if (!dataTables) {
      if (showConsoleMessages) console.log('No tables were found, creating the table and naming it ' + dataWorksheetName);
      let newTable = activeWorkbook.addTable(dataWorksheet.getUsedRange(), true);
      newTable.setName(dataWorksheetName);
    }
    return await exportTableDataToJSON();
  } else {
    if (showConsoleMessages) console.log('Please check the named sheet, it was not found using ' + dataWorksheetName);
  }

  return null
}

async function addLogRecord(startTime: Date, endTime: Date, diff: number, byteSize: number, startsAt: number, endsAt: number, count: number, code: number, message: string): Promise<void> {
    rows.push([startTime.toUTCString(), endTime.toUTCString(), cleanDuration(diff,8), cleanBytes(byteSize), startsAt, endsAt, count, code, message]);
}

async function saveLogRecords(): Promise<object[]> {

  let logWorksheet = activeWorkbook.getWorksheet(logWorksheetName);

  if (logWorksheet) {
    if (showConsoleMessages) console.log('Found the ' + logWorksheetName + ' worksheet');
  } else {
    if (showConsoleMessages) console.log('Adding the log worksheet named ' + logWorksheetName);
    activeWorkbook.addWorksheet(logWorksheetName);
    logWorksheet = activeWorkbook.getWorksheet(logWorksheetName);
  }

  if (logWorksheet != undefined) {

    // Create a header row.
    logWorksheet.getRange('A1:I1').setValues([[LogRecordKeys[0], LogRecordKeys[1], LogRecordKeys[2], LogRecordKeys[3], LogRecordKeys[4], LogRecordKeys[5], LogRecordKeys[6], LogRecordKeys[7], LogRecordKeys[8]]]);

    // Add the data to the specified worksheet, starting at "A2".
    const range = logWorksheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
    range.setValues(rows);

  }

  return null;

}

async function uploadDataToDynatrace(jsonPayload: string, startTime: Date, startsAt: number, endsAt: number, count: number): Promise<object[]> {
  if (showConsoleMessages) console.log('Calling API using ' + apiUrl);
    //console.log('Sending body as ' + jsonPayload);
    //let json: object[];
    //const httpResponse: Promise<Response>;

    const byteSize = byteLengthUTF8(jsonPayload);
    let code: number = 0;
    let jsonObj: Promise<object[]>;
    try {
      await fetch(apiUrl, {
        method: 'POST',
        body: jsonPayload,
        headers: {
          'Content-type': 'application/json',
          'accept': 'application/json; charset=UTF-8',
          'Authorization': apiToken
        }
      })
      .then((response) => {
        if (showConsoleMessages) console.log('API response: ' + response.status);
        code = response.status;
        if (code == 200) {
          jsonObj = response.json();
          return jsonObj;
        } else if (code == 204) {
          return JSON.parse('{ "success": { "code": 204, "message": "Successfully uploaded all data."}}');
        }
      })
      .then((data) => {
        let endTime: Date = new Date();
        let diff = (endTime.getTime() - startTime.getTime()) / (1000 * 3600);

        addLogRecord(startTime, endTime, diff, byteSize, startsAt, endsAt, count, data.success.code, data.success.message);
        return data;
      });
    } catch (error) {
      let endTime: Date = new Date();
      let diff = (endTime.getTime() - startTime.getTime());// / (1000 * 3600);
      console.log(error);
      addLogRecord(startTime, endTime, diff, byteSize, startsAt, endsAt, count, code, error.message);
      return error;
    }        
}


// This function extracts the table data for the 1st table found and calls the convert to JSON function
async function exportTableDataToJSON(): Promise<TableData[]> {
  const dataTable = activeWorkbook.getTable(dataWorksheetName);
    if (dataTable) {
      if (showConsoleMessages) console.log('Found the table');
      if (showConsoleMessages) console.log('Now exporting the table data to JSON');
        const texts = dataTable.getRange().getTexts();
        if (dataTable.getRowCount() > 0) {
          return await returnObjectFromValues(texts);
        }
    } else {
      if (showConsoleMessages) console.log('No tables were found, unable to create the table');
    }
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
async function returnObjectFromValues(values: string[][]): Promise<TableData[]> {

    let objectArray: TableData[] = [];
    let objectKeys: string[] = TableKeys;
    let itemMultiplier = 1;
    let itemCount = 0;
    let startsAt = 0;
    let endsAt = 0;

    if (maxItemCount == 0) {
        maxItemCount = values.length;
    } else {
        maxItemCount;
    }

    startTime = new Date();
    //iterate through the records and prep for api upload
    for (let i = 0; i < maxItemCount; i++) {

        if (i >= itemStartCount && startsAt == 0) {
          startsAt = i;
        }

        if (i >= itemStartCount && values[i] != undefined) {

          itemCount++;
          
          let object: { [key: string]: string } = {};

          //assign object array the key value pairs
          for (let j = 0; j < values[i].length; j++) {
              object[objectKeys[j]] = values[i][j]
          }

          object[objectKeys[values[i].length]] = "SMAXTickets";

          objectArray.push(object as unknown as TableData);

          endsAt = i;
          //use the apiItemCount to send in smaller bundles to api
          if ((i - itemStartCount) == ((apiItemCount * itemMultiplier) - 1) || (i - itemStartCount) == (maxItemCount - 1)) {
            //send these items to api
            let jsonString = JSON.stringify(objectArray);
            if (showConsoleMessages) console.log('Sending ' + itemCount + ' items to Dynatrace');

            let response = await uploadDataToDynatrace(jsonString, startTime, startsAt, endsAt, apiItemCount);


            startTime = new Date();
            itemMultiplier++;
            startsAt = 0;
            endsAt = 0;
            itemCount=0;
            objectArray = [];
            continue;

          }
        }
    }

    if (itemCount > 0) {
      
      //send these items to api
      let jsonString = JSON.stringify(objectArray);
      if (showConsoleMessages) console.log('Sending last items to Dynatrace');

      let response = await uploadDataToDynatrace(jsonString, startTime, startsAt, endsAt, itemCount);

    }

    return objectArray;

}

function cleanDuration(x:number, d:number) {
  return x.toFixed(d) +' ms';
}

function cleanBytes(x: number) {
  let s = x.toString();
  let l = 0, n = parseInt(s, 10) || 0;

  while (n >= 1024 && ++l) {
    n = n / 1024;
  }

  return (n.toFixed(n < 10 && l > 0 ? 1 : 0) + ' ' + units[l]);
}


let testData =
  [
    {
      "content": "Warning: Custom error #1 log sent via Generic Log Ingest",
      "log.source": "/var/log/ingest",
      "timestamp": "2023-11-16T08:02:31.0000",
      "severity": "warn",
      "custom.attribute": "CPU saturation warning",
      "additional.message1": "3 cores",
      "additional.message2": "3 cores"
    },
    {
      "content": "Warning: Custom error #2 log sent via Generic Log Ingest",
      "log.source": "/var/log/ingest",
      "timestamp": "2023-11-16T08:02:41.0000",
      "severity": "warn",
      "custom.attribute": "CPU saturation warning",
      "additional.message1": "3 cores",
      "additional.message2": "3 cores"
    },
    {
      "content": "Exception: Custom error #3 log sent via Generic Log Ingest",
      "log.source": "/var/log/ingest",
      "timestamp": "2023-11-16T08:02:51.0000",
      "severity": "warn",
      "custom.attribute": "CPU saturation warning",
      "additional.message1": "3 cores",
      "additional.message2": "3 cores"
    }
  ];

const TableKeys = [
    "smax.ticketid",
    "smax.timestamp",
    "smax.closedtime",
    "smax.category",
    "smax.solution",
    "content",
    "smax.requestforid",
    "smax.requestedforname",
    "smax.title",
    'smax.requeststatus',
    "smax.completioncode",
    "smax.priority",
    "smax.ownedbyid",
    "smax.ownedbyname",
    "smax..closedbyid",
    "smax.closedbyname",
    "smax.closedbyignid",
    "smax.currentassignment",
    "smax.assignedtogroupid",
    "smax.assignedtogroupname",
    "smax.categoryfirstlevelparent",
    "smax.categorysecondlevelparent",
    "smax.categorytitle",
    "smax.isclosedsameday",
    "log.source"
];

interface TableData {
    "smax.ticketid": string
    "smax.timestamp": string
    "smax.closedtime": string
    "smax.category": number
    "smax.solution": string
    "content": string
    "smax.requestforid": number
    "smax.requestedforname": string
    "smax.title": string
    'smax.requeststatus': string
    "smax.completioncode": string
    "smax.priority": string
    "smax.ownedbyid": number
    "smax.ownedbyname": string
    "smax..closedbyid": number
    "smax.closedbyname": string
    "smax.closedbyignid": number
    "smax.currentassignment": string
    "smax.assignedtogroupid": number
    "smax.assignedtogroupname": string
    "smax.categoryfirstlevelparent": string
    "smax.categorysecondlevelparent": string
    "smax.categorytitle": string
    "smax.isclosedsameday": string
    "log.source": string
};

const LogRecordKeys = [
  "Start Time",
  "End Time",
  "Time Diff",
  "UTF8 Byte Size",
  "Record Starts At",
  "Record Ends At",
  "Items Uploaded",
  "Response Code",
  "Response Message"
];
