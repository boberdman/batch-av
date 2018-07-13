const electrion = require('electron');
const url = require('url');
const path = require('path');
const avatax = require('avatax');

const {app, BrowserWindow} = electron;

let mainWindow;

// Listen for appto be ready
app.on('ready', function(){
  // Create new mainWindow
  mainWindow = new BrowserWindow({});
  // Load html into mainWindow
  mainWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'mainWindow.html'),
    protocol: 'file:',
    slashes: true
  }));
});
