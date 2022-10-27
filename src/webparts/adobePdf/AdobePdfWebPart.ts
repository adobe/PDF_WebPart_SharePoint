export {};
declare global {
  interface Window {
      AdobeDC: any;
  }
}


import { Version } from '@microsoft/sp-core-library';

import PnPTelemetry from "@pnp/telemetry-js";

import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { PropertyFieldPassword } from '@pnp/spfx-property-controls/lib/PropertyFieldPassword';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AdobePdfWebPart.module.scss';
import * as strings from 'AdobePdfWebPartStrings';


export interface IAdobePdfWebPartProps {
  description: string;
  dcclientid: string;
  pdfheight: number;
  pageView: string;
  filePickerResult: IFilePickerResult;
}

export default class AdobePdfWebPart extends BaseClientSideWebPart<IAdobePdfWebPartProps> {


public render(): void {  


 const telemetry = PnPTelemetry.getInstance();
 telemetry.optOut();

 const uniqueDivId:string = "myAdobePDF" + Math.floor(Math.random()*1000) + Date.now();
 const PDFheight:number = this.properties.pdfheight;
 const DCclientid:string = this.properties.dcclientid.substring(0,33);
 const defaultPageView:string = this.properties.pageView;
 let FilePickerResultUrl:string = "";
 let FilePickerResultFile:string = "";

 if ((this.properties.filePickerResult != null) && (DCclientid.length === 32)) {
  FilePickerResultUrl = this.properties.filePickerResult.fileAbsoluteUrl;
  FilePickerResultFile = this.properties.filePickerResult.fileName;
  this.domElement.innerHTML = `

  <div id="` + uniqueDivId + `" style="height: `+ PDFheight + `px; box-shadow: 2px 2px 6px 2px #dadada;"></div>

`;
 }
 else {

  this.domElement.innerHTML = `

  <div class="${ styles.adobePdf }">
  <div class="${ styles.container }">
    <div class="${ styles.row }">
      <div class="${ styles.column }"> <span class="${ styles.title }">`
+ this.properties.description + `</span><br><br><span class="${ styles.description }"><a target="_blank" href="https://documentcloud.adobe.com/dc-integration-creation-app-cdn/main.html?api=pdf-embed-api"> Have you configured a valid Adobe PDF Embed API client ID yet?</a><br> When you're ready, open the web part settings pane to securely paste the client ID.</span> 
</div>
    </div>
  </div>
</div>

  `;
 }
  

// define loadScript function
function loadScript(scriptUrl) {
  let script = document.createElement('script');
  script.src = scriptUrl;
  script.defer = true;
  document.body.appendChild(script);
  
  return new Promise((res, rej) => {
    script.onload  = res;
    script.onerror = rej;
  });
}

// use loadScript
loadScript('https://documentservices.adobe.com/view-sdk/viewer.js')
  .then(() => {
    let adobeDC = window.AdobeDC; 
    if (adobeDC && adobeDC.View) {
      displayPDF();
    } else {
     document.addEventListener("adobe_dc_view_sdk.ready", () => displayPDF());
     }

  })
  .catch(() => {
    console.error('Loading Adobe DC Main script failed.');
  });


 function displayPDF(): void {

   let adobeDC = window.AdobeDC; 
   let adobeDCView = new adobeDC.View({clientId: DCclientid , divId: uniqueDivId});
   adobeDCView.previewFile({
   content:{location: {url: FilePickerResultUrl}},
   metaData:{fileName: FilePickerResultFile}
 }, 
 {defaultViewMode: defaultPageView, showAnnotationTools: false});

 }

}


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Displays a PDF that is stored on this SharePoint site."
          },
          groups: [
            {
              groupName: "Active document",
              groupFields: [
                PropertyFieldFilePicker('filePicker', {
                  context: this.context as any,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { this.properties.filePickerResult = e;  },
                  onChanged: (e: IFilePickerResult) => { this.properties.filePickerResult = e; },
                  key: "filePickerId",
                  buttonLabel: "Browse",
                  label: "PDF documents on this site",
                  hideWebSearchTab:true,
                  hideStockImages:true,
                  hideOneDriveTab:true,
                  hideLocalUploadTab:true,
                  hideLinkUploadTab:true,
                  checkIfFileExists:true,
                  required:true,
                  accepts: ["pdf"]                  
              }), 
              PropertyFieldPassword("dcclientid", {
                key: "clientId",
                label: "Adobe PDF Embed API client ID",
                value: this.properties.dcclientid
              }),
                PropertyPaneSlider('pdfheight',{  
                  label:"Height (in px)",  
                  min:400,  
                  max:700,  
                  value:500,  
                  showValue:true,  
                  step:10                
                }),
                PropertyPaneChoiceGroup('pageView', {
                  label: 'Default page view',
                  options: [
                   { key: 'SINGLE_PAGE', text: 'SINGLE_PAGE', checked: true},
                   { key: 'FIT_PAGE', text: 'FIT_PAGE'},
                   { key: 'FIT_WIDTH', text: 'FIT_WIDTH'},
                   { key: 'TWO_COLUMN', text: 'TWO_COLUMN'}
                 ]
               })
              ]
            }
          ]
        }
      ]
    };
  }
}
