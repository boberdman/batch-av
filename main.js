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
let document;

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
    {
      label: 'Test Credentials',
      click(){
        testAvaTaxCredentials();
      }
    },
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
}}


// Function to open up a file
function openFile() {
   
  var o = dialog.showOpenDialog({ properties: ['openFile'] });
  var workbook = XLSX.readFile(o[0]);
  var worksheet = workbook.Sheets['Sheet1'];
  addressesToValidate = XLSX.utils.sheet_to_json(worksheet);
}

// Call AvaTax to validate an address.
function validateAddress() {
  if (customerAccountNumber == undefined){
    dialog.info('You must set your credentials before testing');
  } else{   
  // Address to be resolved
  // OK, now that you've got this to work for a single row, you'll need to loop through all of the records in the JSON object and write the returned values back to another object and then figure out how to save that object back to a file.
    const address = {
    line1: addressesToValidate[0].line1,
    city: addressesToValidate[0].city,
    postalCode: addressesToValidate[0].postalCode,
    region: addressesToValidate[0].region,
    country: addressesToValidate[0].country
    };
   
    return avaTaxClient.resolveAddress(address)
    .then(result => {
      // address validation result
      console.log(result);
    });
}
}