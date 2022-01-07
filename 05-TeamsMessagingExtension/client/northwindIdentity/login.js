import {
   validateEmployeeLogin,
   setLoggedinEmployeeId
} from './identityService.js';
import {
   getEmployees
} from '../modules/northwindDataService.js';
import { inTeams } from '../modules/teamsHelpers.js';
import 'https://statics.teams.cdn.office.net/sdk/v1.11.0/js/MicrosoftTeams.min.js';

const loginPanel = document.getElementById('loginPanel');
const usernameInput = document.getElementById('username');
const passwordInput = document.getElementById('password');
const loginButton = document.getElementById('loginButton');
const messageDiv = document.getElementById('message');
const hintUL = document.getElementById('hintList');

if (window.location !== window.parent.location) {
   // The page is in an iframe - refuse service
   messageDiv.innerText = "ERROR: You cannot run this app in an IFrame";
} else {

   loginPanel.style.display = 'inline';
   loginButton.addEventListener('click', async ev => {

      messageDiv.innerText = "";
      const employeeId = await validateEmployeeLogin(
         usernameInput.value,
         passwordInput.value
      );
      if (employeeId) {
         setLoggedinEmployeeId(employeeId);
         if (await inTeams()) {
            microsoftTeams.authentication.notifySuccess(employeeId);
         } else {
            window.location.href = document.referrer;
         }
      } else {
         messageDiv.innerText = "Error: user not found";
      }
   });

   (async () => {
      const employees = await getEmployees();

      employees.forEach(employee => {
         const employeeListItem = document.createElement('li');
         employeeListItem.innerHTML = `<b>${employee.lastName.toLowerCase()}</b> (${employee.firstName} ${employee.lastName})`;
         hintUL.appendChild(employeeListItem);
      });
   })();

}