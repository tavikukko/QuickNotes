import {
  BaseClientSideWebPart,
  IWebPartContext,
  IWebPartData
} from '@microsoft/sp-client-preview';
import { DisplayMode } from '@microsoft/sp-client-base';

import { IQuickNotesWebPartProps } from './IQuickNotesWebPartProps';
import importableModuleLoader from '@microsoft/sp-module-loader';

require('./Quicknotes.overrides.css');
require('jquery');
require('adal');

const tenantName: string = "tavikukko365";
const clientId: string = "61287046-f338-481d-a3fd-a2c6a443e32d";
const sharepointApi: string = `https://${tenantName}.sharepoint.com`;
const graphApi: string = "https://graph.microsoft.com";

const authConfig: IAuthenticationConfig = {
  tenant: `${tenantName}.onmicrosoft.com`,
  clientId: `${clientId}`,

  /** where to navigate to after AD logs you out */
  postLogoutRedirectUri: window.location.href,

	/** redirect_uri page, this is the page that receives access tokens
	 *  this URL must match, at least, the scheme and origin of at least 1 of
	 *  the Reply URLs entered on your Azure AD Application configuration page
	 */
  redirectUri: `${window.location.href}`,
  endpoints: {}
  // cacheLocation: "localStorage", // enable this for IE, as sessionStorage does not work for localhost.
};
authConfig.endpoints[sharepointApi] = `https://${tenantName}.sharepoint.com/search`;
authConfig.endpoints[graphApi] = "https://graph.microsoft.com";

export default class QuickNotesWebPart extends BaseClientSideWebPart<IQuickNotesWebPartProps> {

  private Tribute: any;

  public constructor(context: IWebPartContext) {
    super(context);
    importableModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css');
    importableModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');
  }

  private save(editor: any): void {
    this.properties.quicknotecontent = editor.innerHTML;
  }

  public render(mode: DisplayMode = DisplayMode.Read, data?: IWebPartData): void {

    this.domElement.innerHTML = `
      <div id="mentions-wp" contenteditable=true data-placeholder="Type your text here" class="koko">${this.properties.quicknotecontent}</div>
      <hr id="separatpr">
      <button style="margin: 0; display: none;" class="ms-Button ms-Button--primary" id="signinBtn">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Login</span>
        <span class="ms-Button-description">Description of the action this button takes</span>
      </button>
      <button style="margin: 0; display: none;" class="ms-Button" id="signoutBtn">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Logout</span>
        <span class="ms-Button-description"></span>
      </button>
      <span id="loginas"></span>`;

    var editor = document.getElementById('mentions-wp');
    editor.onkeyup = (e) => {
      this.save(editor);
    };

    if (this.displayMode === DisplayMode.Read) {
      document.getElementById('separatpr').style.display = 'none';
      document.getElementById('signinBtn').style.display = 'none';
      document.getElementById('signoutBtn').style.display = 'none';
      document.getElementById('loginas').style.display = 'none';
      document.getElementById('mentions-wp').setAttribute("contenteditable", "false");
    }

    if (this.displayMode === DisplayMode.Edit) {
      importableModuleLoader.loadScript('https://zurb.com/playground/uploads/upload/upload/435/tribute.js', 'Tribute').then((t: any): void => {

        this.Tribute = t;

        var tribute = new this.Tribute({
          menuItemTemplate: item =>
            `<div class="ms-ListItem" style="padding: 0;">
                      <span class="ms-ListItem-secondaryText">`+ item.original.value + `</span>
                      <span class="ms-ListItem-tertiaryText">`+ item.original.email + `</span>
                    </div>`
          ,
          trigger: '#',
          values: [
          ],
          selectTemplate: item =>
            '<a class="ms-Link" href="' + item.original.email + '" target="_blank" title="' + item.original.email + '">' + item.original.value + '</a>'
        });

        var tribute2 = new this.Tribute({
          menuItemTemplate: item =>
            `<div tabindex="0" role="button" class="ms-PeoplePicker-peopleListBtn">
                      <div class="ms-Persona ms-Persona--selectable ms-Persona--sm">
                        <div class="ms-Persona-imageArea">
                          <div class="ms-Persona-initials ms-Persona-initials--darkBlue">TT</div>
                        </div>
                        <div class="ms-Persona-details">
                          <div class="ms-Persona-primaryText">`+ item.original.value + `</div>
                          <div class="ms-Persona-secondaryText">`+ item.original.email + `</div>
                        </div>
                      </div>
                    </div>`,
          trigger: '@',
          values: [
          ],
          selectTemplate: item =>
            '<a class="ms-Link" href="mailto:' + item.original.email + '" target="_top" title="' + item.original.value + '">' + item.original.value + '</a>'
        });

        tribute.attach(document.getElementById('mentions-wp'));
        tribute2.attach(document.getElementById('mentions-wp'));

        document.getElementById('mentions-wp').addEventListener('tribute-no-match', function (e) {

          const authContext: AuthenticationContext = new AuthenticationContext(authConfig);
          //const graphApi: string = "https://graph.microsoft.com";

          const self: any = this;
          const d: any = jQuery.Deferred<any>();

          authContext.acquireToken(graphApi, (error: string, token: string) => {
            if (error || !token) {
              const msg: any = `ADAL error occurred: ${error}`;
              d.rejectWith(this, [msg]);
              return;
            }
            let url = `${graphApi}/v1.0/me/drive/root/children?$filter=startswith(name,'${tribute.current.mentionText}')`;
            if (tribute2.isActive)
              url = `${graphApi}/v1.0/users?$filter=startswith(userPrincipalName,'${tribute2.current.mentionText}') or startswith(mail,'${tribute2.current.mentionText}') or startswith(surname,'${tribute2.current.mentionText}')`;

            jQuery.ajax({
              type: "GET",
              url: url,
              headers: {
                "Accept": "application/json;odata.metadata=minimal",
                "Authorization": `Bearer ${token}`
              }
            }).done((response: { value: any[] }) => {
              console.log("Successfully fetched data from O365.");
              d.resolveWith(self);
              if (response["@odata.context"].includes("/drive/root/children")) {
                var documents: any = response.value;
                for (var document of documents) {
                  if (tribute.collection[0].values.filter( d => d.key.toLowerCase() == document.name.toLowerCase()).length == 0)
                    tribute.append(0, [
                      { key: document.name, value: document.name, email: document.webUrl }
                    ]);
                }
              }
              else {
                var users: any = response.value;
                for (var user of users) {
                  let email = '';
                  if (user.email == undefined) email = user.userPrincipalName.toLowerCase();
                  else email = user.mail.toLowerCase();

                  if (tribute2.collection[0].values.filter(u => u.key.toLowerCase() == email).length == 0) {
                    tribute2.append(0, [
                      { key: email, value: user.displayName, email: email }
                    ]);
                  }
                }
              }
            }).fail((xhr: JQueryXHR) => {
              const msg: any = `Fetching data from Office365 failed. ${xhr.status}: ${xhr.statusText}`;
              console.log(msg);
              d.rejectWith(self, [msg]);
            });
          });
        });

      });
    }
    if (this.displayMode === DisplayMode.Edit) this.manageAuthentication();

  }

  private manageAuthentication(): void {

    const authContext: AuthenticationContext = new AuthenticationContext(authConfig);

    const isCallback: any = authContext.isCallback(window.location.hash);
    if (isCallback) {
      const loginReq: any = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
      console.log(`IS Callback! ${loginReq}`);
      authContext.handleWindowCallback();
      return;
    }
    console.log("Is NOT Callback!");

    var user: IUserInfo = authContext.getCachedUser();

    if (user) {
      console.log(`User is logged-in: ${JSON.stringify(user)}`);
      document.getElementById('loginas').innerHTML = '<i class="ms-Icon ms-Icon--person" aria-hidden="true"></i> ' + user.userName;

      document.getElementById('signinBtn').style.display = 'none';
      document.getElementById('signoutBtn').style.display = '';

      document.getElementById('signoutBtn').onclick = () => {
        authContext.logOut();
      };
    } else {
      console.log("User is NOT logged-in!!");
      authContext.clearCache();
      document.getElementById('signoutBtn').style.display = 'none';
      document.getElementById('signinBtn').style.display = '';
      document.getElementById('signinBtn').onclick = () => {
        authContext.login();
      };
    }

  }

}
