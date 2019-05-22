import { Component } from '@angular/core';
const template = require('./app.component.html');

@Component({
  selector: 'app-home',
  template
})
export default class AppComponent {
  welcomeMessage = 'Welcome';

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

  addImage() {
    FileReader.readFileSync()
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
}
