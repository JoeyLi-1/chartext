import { Component } from '@angular/core';
const template = require('./app.component.html');
// import { HttpClient } from "@angular/common/http";

declare let Auth0: any;
declare let $: any;

@Component({
  selector: 'app-home',
  template
})
export default class AppComponent {
  welcomeMessage = 'Welcome';
  auth0: any;
  // debugData = 'nothing';
  // constructor(private http: HttpClient){

  // }

  async run() {
    /**
   * Insert your PowerPoint code here
   */
    Office.context.document.setSelectedDataAsync("Hello World!",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }

  addText() {
    Office.context.document.setSelectedDataAsync('Hello World!',
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                this.showNotification("Error", asyncResult.error.message);
            }
        });
  }

  showNotification(header, content) {
    
  }

  async addImage() {
    // https://docs.microsoft.com/en-us/office/dev/add-ins/images/office-add-ins-my-account.png
    // const data = await this.http.get("https://docs.microsoft.com/en-us/office/dev/add-ins/images/office-add-ins-my-account.png").toPromise();
    // this.insertImageFromBase64String(data);
    let that = this;
    $.ajax({
        url: "/assets/logo.png", success: function (result) {
          that.insertImageFromBase64String(result);
        }, error: function (xhr, status, error) {
            // showNotification("Error", "Oops, something went wrong.");
            console.log(error);
        }
    });
  }

  insertImageFromBase64String(image) {
    // Call Office.js to insert the image into the document.
    Office.context.document.setSelectedDataAsync(image, {
        coercionType: Office.CoercionType.Image
    },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                this.showNotification("Error", asyncResult.error.message);
            }
        });
  }

  login() {
    // this.auth0.client.login({
    //   realm: 'Username-Password-Authentication', //connection name or HRD domain
    //   username: 'joey@clarifyhealth.com',
    //   password: 'Clarify1',
    //   audience: '',
    //   scope: 'openid name email',
    //   }, function(err, authResult) {
    //     this.debugData = authResult.toString();
    //     console.table(authResult);
    // });
    this.auth0 = new Auth0({
        domain: 'clarifyhealth.auth0.com',
        clientID: 'loXgNdZtCNuf8fe8J9qeAtyFTeXbFlJn',
        callbackURL: '',
        responseType: 'token'
    });
    try {
        this.auth0.login({
            connection: 'Username-Password-Authentication',
            email: 'joey@clarifyhealth.com',
            password: 'Clarify1',
            responseType: 'token',
            sso: false,
            scope: 'openid name email app_metadata identities'
        }, (err: any, result: any) => {
          console.log(result);
        });
    }
    catch (err) {
      console.log(err);
    }
  }
}
