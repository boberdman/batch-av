// Load Required Modules
// Electron framework
// URL for url construction? #TODO: look up the url function
// Path to check out the local file path? #TODO: look up the path function
// AvaTax for the AvaTax SDK
// Remote to load the dialogs component of the os
// FS = load the file systems to execute our common tasks (CRUD)
const electron = require('electron');
const url = require('url');
const path = require('path');
const avaTax = require('avatax');
const XLSX = require('xlsx');
const {app, BrowserWindow, Menu, ipcMain} = electron;
const {dialog} = require('electron'); 


// Set AvaTax Credentials 
let customerAccountNumber;
let customerSoftwareLicenseKey;
var validatedAddressResults = [];


const avaTaxConfig = {
  appName:"Bob's Batch Address Validator",
  appVersion: '1.0',
  environment: 'production',
  machineName: 'hal-1000'
};

// Set Window Variables
let mainWindow;
let addWindow;

// Listen for appto be ready
app.on('ready', function(){
  // Create new Window
  mainWindow = new BrowserWindow({});
  // Load html into mainWindow
  mainWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'mainWindow.html'),
    protocol: 'file:',
    slashes: true
  }));
  // Quit app when closed
  mainWindow.on('closed', function(){
    app.quit();
  });

  // Build Menu from mainMenuTemplate
  const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
  // Insert Menu
  Menu.setApplicationMenu(mainMenu);
});


// Handle create add window
function createAddWindow(){
  // Create new Window
  addWindow = new BrowserWindow({
    width: 300,
    height: 200,
    title: 'Authentication'
  });
  // Load html into mainWindow
  addWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'addWindow.html'),
    protocol: 'file:',
    slashes: true
  }));
  // Garbage collection handle
  addWindow.on('close', function(){
    addWindow = null;
  });
}

// Catch accountNumber:add
ipcMain.on('accountNumber:add',function(e, accountNumber){
  //console.log(accountNumber);
  mainWindow.webContents.send('accountNumber:add', accountNumber);
  customerAccountNumber = accountNumber;
});

// Catch softwareLicenseKey:add
ipcMain.on('softwareLicenseKey:add',function(e, softwareLicenseKey){
  //console.log(softwareLicenseKey);
  mainWindow.webContents.send('softwareLicenseKey:add', softwareLicenseKey);
  addWindow.close();
  customerSoftwareLicenseKey = softwareLicenseKey;
  // Make AvaTax Credentials
  const avaTaxCreds = {
    username: customerAccountNumber,
    password: customerSoftwareLicenseKey
  };
  // Make the AvaTax Client
  avaTaxClient = new avaTax(avaTaxConfig).withSecurity(avaTaxCreds);
});


// Create menu template
const mainMenuTemplate = [
  {
  label: 'File',
  submenu: [
    {
      label: 'Authentication',
      click(){
        createAddWindow();
      }
    },
    // {
    //   label: 'Test Credentials',
    //   click(){
    //     testAvaTaxCredentials();
    //   }
    // },
    {
      label: 'Open File',
      accelerator: process.platform == 'darwin' ? 'Command+O' : 'Ctrl+O',
      click(){
        openFile();
      }
    },
    {
      label: 'Validate Address', 
      click(){
        validateAddress();
      }
    },
    {
      label: 'Save File', 
      click(){
        saveFile();
      }
    },
    {
      label: 'Quit',
      accelerator: process.platform == 'darwin' ? 'Command+Q' : 'Ctrl+Q',
      click(){
        app.quit();
      }
    }
  ]
}
];

// If mac, add empty object to menu
if(process.platform == 'darwin'){
  mainMenuTemplate.unshift({});
}

// Add developer tools item if not in prodcution
if(process.env.NODE_ENV !== 'production'){
  mainMenuTemplate.push({
    label: 'Developer Tools',
    submenu: [
      {
      label: 'Toggle DevTools',
      accelerator: process.platform == 'darwin' ? 'Command+I' : 'Ctrl+I',
      click(item, focusedWindow){
        focusedWindow.toggleDevTools();
      }
      },
      {
        role: 'reload'
      }
    ]
  });
}

// test logging the AvaTax credentials:
function testAvaTaxCredentials() {
  if (customerAccountNumber == undefined){
    dialog.info('You must set your credentials before testing');
  } else{
  console.log(customerAccountNumber);
  console.log(customerSoftwareLicenseKey);
}
}


// Function to open up a file
function openFile() {
   
  var o = dialog.showOpenDialog({ properties: ['openFile'] });
  var workbook = XLSX.readFile(o[0]);
  worksheet = workbook.Sheets['Sheet1'];
  addressesToValidate = XLSX.utils.sheet_to_json(worksheet);
  console.log(addressesToValidate);
}
//function to count JSON array length, number of returned results in the address validation
function objectLength(obj) {
  var result = 0;
  for(var prop in obj) {
    if (obj.hasOwnProperty(prop)) {
    // or Object.prototype.hasOwnProperty.call(obj, prop)
      result++;
    }
  }
  return result;
}

// Call AvaTax to validate an address.
function validateAddress() {

  if (customerAccountNumber == undefined){
    dialog.info('Whoa buddy! You must set your credentials before testing');
  } else{   
  // Address to be resolved
  for (var i=0, len =  objectLength(addressesToValidate); i < len; i++){
    var address = {
    line1: addressesToValidate[i].line1,
    city: addressesToValidate[i].city,
    postalCode: addressesToValidate[i].postalCode,
    region: addressesToValidate[i].region,
    country: addressesToValidate[i].country
    };
    // Call Avalara to validate the address
    avaTaxClient.resolveAddress(address)
    .then(result => {
      // address validation result
       if (result === undefined || result.resolutionQuality == 0 || result.resolutionQuality == undefined){
        //resultsWorksheet = XLSX.utils.json_to_sheet(validatedAddressResults, {header: resultValidatedAddressResults.keys});
        dialog.showErrorBox("whoops! someting went wrong")
      } else {
        validatedAddressResults.push(result);
   
      };
    });
  };
};
}

// Save Validated Addresses
function saveFile(){
  console.log(validatedAddressResults);
  if (validatedAddressResults == null || validatedAddressResults == undefined){
    dialog.showErrorBox('Nothing to Save','try validating some addresses first')
  } else {
    /* show a file-save dialog and write the workbook */
    // Create New Workbook
    /* make the worksheet */
    var ws = XLSX.utils.aoa_to_sheet(validatedAddressResults);
      console.log("this is the worksheet output")
      console.log(ws);

      /* add worksheet to workbook */
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Validated_Results");
      console.log("this is the workbook output");
      console.log(wb);
   
      var o = dialog.showSaveDialog();
      XLSX.writeFile(wb, o);  };
}