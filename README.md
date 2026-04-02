# Adobe PDF Viewer web part for Microsoft SharePoint Online
## Modernized for MSFT SPFx 1.22.x / Node.js 22 LTS / TypeScript 5.8 / Heft.
## No third-party dependencies.

## Summary

This project contains sample code that implements the core capabilities of the Adobe PDF Embed API for use in a Microsoft SharePoint Framework (SPFx) web part. The Adobe PDF Viewer web part leverages Adobe's PDF viewer. It has been built for use in "site pages" in Microsoft SharePoint Online, and it can render PDF's stored within SharePoint document libraries without losing SharePoint site navigation. Users can easily print or download a PDF from the web part. The default view mode is configurable in the web part settings. 

![Image of a cat in a PDF document displayed on a SharePoint Online site page](PDF-demo-screenshot.png "Screenshot of PDF Viewer web part on SharePoint Online site page")

## Used SharePoint Framework Version

SPFx 1.22.2

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Adobe PDF Embed](https://developer.adobe.com/document-services/apis/pdf-embed)
- [Adobe UI customization for PDF's](https://developer.adobe.com/document-services/docs/overview/pdf-embed-api/howtos_ui/)


## Prerequisites

The PDF Viewer web part requires an Adobe client ID which can be created by following the steps here: https://documentcloud.adobe.com/dc-integration-creation-app-cdn/main.html?api=pdf-embed-api. Configure the domain associated with the Adobe client ID as sharepoint.com or MSTENANTNAME.sharepoint.com. 

In addition, the URLs https://documentservices.adobe.com and https://acrobatservices.adobe.com need to be added as trusted script sources in Microsoft's SharePoint Online Admin Center. 

This project was built for SPFx 1.22 or newer, and it was tested with Node.js version 22.22.2.


## Recommendations for use in production

- Customize the web part code to meet your production requirements.
- Customize the web part code to pull the Adobe Embed client ID from a pre-defined, configurable location, e.g. a SharePoint list, in the same Microsoft 365 tenant.


## License

**This project is licensed under the terms of the MIT license.**

Copyright 2022-2026 Adobe

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

---

## How to build

- Clone this repository
- Ensure that you are at the project folder
- In the command line run:
  - **npm install**
  - **npm run build**

During the process it may also be helpful to run the following:
  - npm audit fix (do not use the --force switch)

To become familiar with Microsoft's SPFx and Heft, you can review the following Microsoft articles:
  1) https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-toolchain-rushstack-heft
  2) https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part



