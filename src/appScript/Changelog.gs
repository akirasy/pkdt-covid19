function aboutLicense() {
  var title = 'Open Source';
  var subtitle = `
    This app is open source and free to use under the terms of GNU General Public License v3.0.

    Copyright (C) 2021  akirasy
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}

function aboutAuthor() {
  var title = 'AppScript Author';
  var subtitle = `
    This app is developed by akirasy <fitri.abakar@gmail.com>
    
    Feel free to browse other app here --> https://github.com/akirasy
    For this specific app source, look here --> https://gist.github.com/akirasy/f249851b524b0da82ed9b62f5edb5639
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}

function aboutGoogleAppScript() {
  var title = 'Google AppScript';
  var subtitle = `
    Google Apps Script is a rapid application development platform that makes it 
    fast and easy to create business applications that integrate with Google Workspace. 
    
    You write code in modern JavaScript and have access to built-in libraries for favorite 
    Google Workspace applications like Gmail, Calendar, Drive, and more. 
    
    There's nothing to install—we give you a code editor right in your browser, 
    and your scripts run on Google's servers.

    Learn more at --> https://developers.google.com/apps-script/overview
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}
