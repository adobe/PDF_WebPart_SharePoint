"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_loader_1 = require("@microsoft/sp-loader");
var PropertyPaneFilePicker_1 = require("./propertyPane/PropertyPaneFilePicker");
var AdobePdfWebPart_module_scss_1 = tslib_1.__importDefault(require("./AdobePdfWebPart.module.scss"));
var strings = tslib_1.__importStar(require("AdobePdfWebPartStrings"));
/* ------------------------------------------------------------------ */
/*  Web Part                                                           */
/* ------------------------------------------------------------------ */
var AdobePdfWebPart = /** @class */ (function (_super) {
    tslib_1.__extends(AdobePdfWebPart, _super);
    function AdobePdfWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /* ---- lifecycle ------------------------------------------------- */
    AdobePdfWebPart.prototype.render = function () {
        var _this = this;
        var _a;
        this.domElement.innerHTML = this._buildHtml();
        if (this.properties.clientId && ((_a = this.properties.filePickerResult) === null || _a === void 0 ? void 0 : _a.fileAbsoluteUrl)) {
            this._loadAdobeSdk()
                .then(function () { return _this._renderPdf(); })
                .catch(function (err) {
                console.error('[AdobePdfWebPart] SDK load error:', err);
                _this._showError(strings.SdkLoadError);
            });
        }
    };
    Object.defineProperty(AdobePdfWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    /* ---- property pane -------------------------------------------- */
    AdobePdfWebPart.prototype.getPropertyPaneConfiguration = function () {
        var _this = this;
        var _a;
        return {
            pages: [
                {
                    header: { description: strings.PropertyPaneDescription },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                {
                                    type: sp_property_pane_1.PropertyPaneFieldType.Custom,
                                    targetProperty: 'clientId',
                                    properties: {
                                        key: 'clientIdField',
                                        onRender: function (elem) {
                                            elem.innerHTML = "\n                        <label style=\"display:block;font-weight:600;font-size:14px;padding-bottom:5px;font-family:'Segoe UI',sans-serif\">\n                          ".concat(strings.ClientIdFieldLabel, "\n                        </label>\n                        <input type=\"password\" value=\"").concat(_this.properties.clientId || '', "\"\n                          style=\"width:100%;padding:6px 8px;border:1px solid #8a8886;border-radius:4px;font-size:13px;font-family:'Segoe UI',sans-serif;box-sizing:border-box\"\n                          autocomplete=\"off\" />\n                        <p style=\"font-size:11px;color:#605e5c;margin:4px 0 0;font-family:'Segoe UI',sans-serif\">\n                          ").concat(strings.ClientIdFieldDescription, "\n                        </p>");
                                            elem.querySelector('input').addEventListener('input', function (e) {
                                                _this.properties.clientId = e.target.value;
                                                _this.render();
                                            });
                                        },
                                        onDispose: function () { }
                                    }
                                },
                                (0, PropertyPaneFilePicker_1.PropertyPaneFilePicker)('filePickerResult', {
                                    key: 'filePickerId',
                                    label: strings.FilePickerLabel,
                                    buttonLabel: strings.FilePickerButtonLabel,
                                    accepts: ['.pdf'],
                                    value: this.properties.filePickerResult,
                                    webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
                                    spHttpClient: this.context.spHttpClient,
                                    onSelect: function (result) {
                                        _this.properties.filePickerResult = result;
                                        _this.render();
                                        // Force property pane to re-render so it shows the new filename
                                        _this.context.propertyPane.refresh();
                                    }
                                }),
                                (0, sp_property_pane_1.PropertyPaneDropdown)('viewMode', {
                                    label: strings.ViewModeFieldLabel,
                                    options: AdobePdfWebPart.EMBED_MODES,
                                    selectedKey: (_a = this.properties.viewMode) !== null && _a !== void 0 ? _a : 'FULL_WINDOW'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    /* ---- private helpers ------------------------------------------ */
    AdobePdfWebPart.prototype._buildHtml = function () {
        var _a, _b, _c;
        var fileName = (_b = (_a = this.properties.filePickerResult) === null || _a === void 0 ? void 0 : _a.fileName) !== null && _b !== void 0 ? _b : '';
        var hasFile = !!((_c = this.properties.filePickerResult) === null || _c === void 0 ? void 0 : _c.fileAbsoluteUrl);
        var hasClientId = !!this.properties.clientId;
        if (!hasClientId) {
            return "\n        <div class=\"".concat(AdobePdfWebPart_module_scss_1.default.adobePdf, "\">\n          <div class=\"").concat(AdobePdfWebPart_module_scss_1.default.container, "\">\n            <p class=\"").concat(AdobePdfWebPart_module_scss_1.default.message, "\">").concat(strings.MissingClientId, "</p>\n          </div>\n        </div>");
        }
        if (!hasFile) {
            return "\n        <div class=\"".concat(AdobePdfWebPart_module_scss_1.default.adobePdf, "\">\n          <div class=\"").concat(AdobePdfWebPart_module_scss_1.default.container, "\">\n            <p class=\"").concat(AdobePdfWebPart_module_scss_1.default.message, "\">").concat(strings.MissingFile, "</p>\n          </div>\n        </div>");
        }
        return "\n      <div class=\"".concat(AdobePdfWebPart_module_scss_1.default.adobePdf, "\">\n        <div class=\"").concat(AdobePdfWebPart_module_scss_1.default.header, "\">\n          <span class=\"").concat(AdobePdfWebPart_module_scss_1.default.fileName, "\">").concat(this._escapeHtml(fileName), "</span>\n        </div>\n        <div id=\"").concat(AdobePdfWebPart.VIEWER_DIV_ID, "\" class=\"").concat(AdobePdfWebPart_module_scss_1.default.viewer, "\"></div>\n      </div>");
    };
    AdobePdfWebPart.prototype._loadAdobeSdk = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if ((_a = window.AdobeDC) === null || _a === void 0 ? void 0 : _a.View) {
                            return [2 /*return*/, Promise.resolve()];
                        }
                        return [4 /*yield*/, sp_loader_1.SPComponentLoader.loadScript(AdobePdfWebPart.ADOBE_SDK_URL, { globalExportsName: 'AdobeDC' })];
                    case 1:
                        _b.sent();
                        return [4 /*yield*/, new Promise(function (resolve) {
                                var _a;
                                if ((_a = window.AdobeDC) === null || _a === void 0 ? void 0 : _a.View) {
                                    resolve();
                                }
                                else {
                                    document.addEventListener('adobe_dc_view_sdk.ready', function () { return resolve(); }, { once: true });
                                }
                            })];
                    case 2: return [2 /*return*/, _b.sent()];
                }
            });
        });
    };
    AdobePdfWebPart.prototype._renderPdf = function () {
        var _a, _b, _c, _d, _e;
        var fileUrl = (_a = this.properties.filePickerResult) === null || _a === void 0 ? void 0 : _a.fileAbsoluteUrl;
        var fileName = (_c = (_b = this.properties.filePickerResult) === null || _b === void 0 ? void 0 : _b.fileName) !== null && _c !== void 0 ? _c : 'document.pdf';
        if (!fileUrl || !((_d = window.AdobeDC) === null || _d === void 0 ? void 0 : _d.View))
            return;
        var downloadUrl = fileUrl;
        var adobeDCView = new window.AdobeDC.View({
            clientId: this.properties.clientId,
            divId: AdobePdfWebPart.VIEWER_DIV_ID
        });
        adobeDCView.previewFile({
            content: { location: { url: downloadUrl } },
            metaData: { fileName: fileName }
        }, {
            embedMode: (_e = this.properties.viewMode) !== null && _e !== void 0 ? _e : 'FULL_WINDOW',
            showDownloadPDF: true,
            showPrintPDF: true
        });
    };
    AdobePdfWebPart.prototype._showError = function (message) {
        var viewer = this.domElement.querySelector("#".concat(AdobePdfWebPart.VIEWER_DIV_ID));
        if (viewer) {
            viewer.innerHTML = "<p class=\"".concat(AdobePdfWebPart_module_scss_1.default.message, "\">").concat(this._escapeHtml(message), "</p>");
        }
    };
    AdobePdfWebPart.prototype._escapeHtml = function (text) {
        var map = {
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, function (ch) { var _a; return (_a = map[ch]) !== null && _a !== void 0 ? _a : ch; });
    };
    AdobePdfWebPart.ADOBE_SDK_URL = 'https://acrobatservices.adobe.com/view-sdk/viewer.js';
    AdobePdfWebPart.EMBED_MODES = [
        { key: 'FULL_WINDOW', text: 'Full Window' },
        { key: 'SIZED_CONTAINER', text: 'Sized Container' },
        { key: 'IN_LINE', text: 'In-Line' },
        { key: 'LIGHT_BOX', text: 'Light Box' }
    ];
    AdobePdfWebPart.VIEWER_DIV_ID = 'adobe-dc-view';
    return AdobePdfWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = AdobePdfWebPart;
//# sourceMappingURL=AdobePdfWebPart.js.map