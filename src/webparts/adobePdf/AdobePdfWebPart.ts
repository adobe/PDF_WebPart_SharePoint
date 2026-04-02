/**
 * AdobePdfWebPart.ts
 *
 * Adobe PDF Viewer web part for Microsoft SharePoint Online.
 * Modernized for SPFx 1.22.x / Node.js 22 LTS / TypeScript 5.8 / Heft.
 *
 * This version has no third-party runtime dependencies, the property
 * pane file picker is implemented natively using SPHttpClient + DOM.
 *
 * Copyright 2022-2026 Adobe – MIT License
 */

import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  PropertyPaneFilePicker,
  type IFilePickerResult
} from './propertyPane/PropertyPaneFilePicker';

import styles from './AdobePdfWebPart.module.scss';
import * as strings from 'AdobePdfWebPartStrings';

/* ------------------------------------------------------------------ */
/*  Interfaces                                                         */
/* ------------------------------------------------------------------ */

export interface IAdobePdfWebPartProps {
  /** Adobe PDF Embed API client ID */
  clientId: string;
  /** Selected PDF file result */
  filePickerResult: IFilePickerResult | undefined;
  /** Adobe Embed view mode */
  viewMode: string;
}

/** Adobe DC View SDK types */
interface IAdobeDCView {
  previewFile(
      content: { content: { location: { url:string}}; metaData: { fileName: string }},
      options: { embedMode: string; showDownloadPDF: boolean; showPrintPDF: boolean }
  ): void;
}


interface IAdobeDCViewConstructor {
  new (config: { clientId: string; divId: string }): IAdobeDCView;
}

declare global {
  interface Window {
    AdobeDC?: { View?: IAdobeDCViewConstructor };
  }
}

/* ------------------------------------------------------------------ */
/*  Web Part                                                           */
/* ------------------------------------------------------------------ */

export default class AdobePdfWebPart extends BaseClientSideWebPart<IAdobePdfWebPartProps> {

  private static readonly ADOBE_SDK_URL =
    'https://acrobatservices.adobe.com/view-sdk/viewer.js';

  private static readonly EMBED_MODES: IPropertyPaneDropdownOption[] = [
    { key: 'FULL_WINDOW',     text: 'Full Window' },
    { key: 'SIZED_CONTAINER', text: 'Sized Container' },
    { key: 'IN_LINE',         text: 'In-Line' },
    { key: 'LIGHT_BOX',       text: 'Light Box' }
  ];

  private static readonly VIEWER_DIV_ID = 'adobe-dc-view';

  /* ---- lifecycle ------------------------------------------------- */

  public render(): void {
    this.domElement.innerHTML = this._buildHtml();

    if (this.properties.clientId && this.properties.filePickerResult?.fileAbsoluteUrl) {
      this._loadAdobeSdk()
        .then(() => this._renderPdf())
        .catch((err: Error) => {
          console.error('[AdobePdfWebPart] SDK load error:', err);
          this._showError(strings.SdkLoadError);
        });
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /* ---- property pane -------------------------------------------- */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                {
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'clientId',
                  properties: {
                    key: 'clientIdField',
                    onRender: (elem: HTMLElement) => {
                      elem.innerHTML = `
                        <label style="display:block;font-weight:600;font-size:14px;padding-bottom:5px;font-family:'Segoe UI',sans-serif">
                          ${strings.ClientIdFieldLabel}
                        </label>
                        <input type="password" value="${this.properties.clientId || ''}"
                          style="width:100%;padding:6px 8px;border:1px solid #8a8886;border-radius:4px;font-size:13px;font-family:'Segoe UI',sans-serif;box-sizing:border-box"
                          autocomplete="off" />
                        <p style="font-size:11px;color:#605e5c;margin:4px 0 0;font-family:'Segoe UI',sans-serif">
                          ${strings.ClientIdFieldDescription}
                        </p>`;
                      elem.querySelector('input')!.addEventListener('input', (e) => {
                        this.properties.clientId = (e.target as HTMLInputElement).value;
                        this.render();
                      });
                    },
                    onDispose: () => {}
                  }
                } as IPropertyPaneField<IPropertyPaneCustomFieldProps>,
                PropertyPaneFilePicker('filePickerResult', {
                  key: 'filePickerId',
                  label: strings.FilePickerLabel,
                  buttonLabel: strings.FilePickerButtonLabel,
                  accepts: ['.pdf'],
                  value: this.properties.filePickerResult,
                  webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
                  spHttpClient: this.context.spHttpClient,
                  onSelect: (result: IFilePickerResult) => {
                    this.properties.filePickerResult = result;
                    this.render();
                    // Force property pane to re-render so it shows the new filename
                    this.context.propertyPane.refresh();
                  }
                }),
                PropertyPaneDropdown('viewMode', {
                  label: strings.ViewModeFieldLabel,
                  options: AdobePdfWebPart.EMBED_MODES,
                  selectedKey: this.properties.viewMode ?? 'FULL_WINDOW'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /* ---- private helpers ------------------------------------------ */

  private _buildHtml(): string {
    const fileName = this.properties.filePickerResult?.fileName ?? '';
    const hasFile = !!this.properties.filePickerResult?.fileAbsoluteUrl;
    const hasClientId = !!this.properties.clientId;

    if (!hasClientId) {
      return `
        <div class="${styles.adobePdf}">
          <div class="${styles.container}">
            <p class="${styles.message}">${strings.MissingClientId}</p>
          </div>
        </div>`;
    }

    if (!hasFile) {
      return `
        <div class="${styles.adobePdf}">
          <div class="${styles.container}">
            <p class="${styles.message}">${strings.MissingFile}</p>
          </div>
        </div>`;
    }

    return `
      <div class="${styles.adobePdf}">
        <div class="${styles.header}">
          <span class="${styles.fileName}">${this._escapeHtml(fileName)}</span>
        </div>
        <div id="${AdobePdfWebPart.VIEWER_DIV_ID}" class="${styles.viewer}"></div>
      </div>`;
  }

 private async _loadAdobeSdk(): Promise<void> {
    if (window.AdobeDC?.View) {
      return Promise.resolve();
    }

    await SPComponentLoader.loadScript(
     AdobePdfWebPart.ADOBE_SDK_URL,
     { globalExportsName: 'AdobeDC' }
   );
   return await new Promise<void>((resolve) => {
     if (window.AdobeDC?.View) {
       resolve();
     } else {
       document.addEventListener('adobe_dc_view_sdk.ready', () => resolve(), { once: true });
     }
   });
  }


  private _renderPdf(): void {
    const fileUrl = this.properties.filePickerResult?.fileAbsoluteUrl;
    const fileName = this.properties.filePickerResult?.fileName ?? 'document.pdf';

    if (!fileUrl || !window.AdobeDC?.View) return;

    const downloadUrl = fileUrl;

    const adobeDCView = new window.AdobeDC.View({
      clientId: this.properties.clientId,
      divId: AdobePdfWebPart.VIEWER_DIV_ID
    });

    adobeDCView.previewFile(
      {
        content:  { location: { url: downloadUrl } },
        metaData: { fileName }
      },
      {
        embedMode:       this.properties.viewMode ?? 'FULL_WINDOW',
        showDownloadPDF: true,
        showPrintPDF:    true
      }
    );
  }

  private _showError(message: string): void {
    const viewer = this.domElement.querySelector(`#${AdobePdfWebPart.VIEWER_DIV_ID}`);
    if (viewer) {
      viewer.innerHTML = `<p class="${styles.message}">${this._escapeHtml(message)}</p>`;
    }
  }

  private _escapeHtml(text: string): string {
    const map: Record<string, string> = {
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, (ch) => map[ch] ?? ch);
  }
}
