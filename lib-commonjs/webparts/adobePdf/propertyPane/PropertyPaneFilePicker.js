"use strict";
/**
 * PropertyPaneFilePicker.ts
 *
 * A zero-dependency, vanilla-TypeScript replacement for the PnP
 * PropertyFieldFilePicker. Renders a "Select PDF" button in the
 * property pane; when clicked, opens a modal panel that lets the
 * user browse document libraries on the current site and pick a
 * PDF file.
 *
 * Uses only:
 *  - @microsoft/sp-property-pane  (IPropertyPaneField)
 *  - @microsoft/sp-http           (SPHttpClient)
 *  - Plain DOM APIs
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.PropertyPaneFilePicker = PropertyPaneFilePicker;
var tslib_1 = require("tslib");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_http_1 = require("@microsoft/sp-http");
/* ------------------------------------------------------------------ */
/*  Factory function                                                   */
/* ------------------------------------------------------------------ */
function PropertyPaneFilePicker(targetProperty, props) {
    return {
        type: sp_property_pane_1.PropertyPaneFieldType.Custom,
        targetProperty: targetProperty,
        properties: {
            key: props.key,
            onRender: function (domElement) {
                renderField(domElement, props);
            },
            onDispose: function (domElement) {
                domElement.innerHTML = '';
            }
        }
    };
}
/* ------------------------------------------------------------------ */
/*  Property-pane field renderer                                       */
/* ------------------------------------------------------------------ */
function renderField(container, props) {
    var _a, _b;
    container.innerHTML = '';
    // -- Label
    var label = el('label', {
        style: 'display:block;font-weight:600;font-size:14px;padding-bottom:5px;color:var(--spPageText,#323130);font-family:"Segoe UI",sans-serif'
    }, props.label);
    container.appendChild(label);
    // -- Selected file display
    var fileName = (_b = (_a = props.value) === null || _a === void 0 ? void 0 : _a.fileName) !== null && _b !== void 0 ? _b : '';
    if (fileName) {
        var selected = el('div', {
            style: 'display:flex;align-items:center;gap:6px;padding:6px 0 8px 0;font-size:13px;color:var(--spPageText,#605e5c);font-family:"Segoe UI",sans-serif'
        });
        selected.appendChild(el('span', { style: 'font-size:15px' }, '📄'));
        selected.appendChild(el('span', {
            style: 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:200px',
            title: fileName
        }, fileName));
        container.appendChild(selected);
    }
    // -- Browse button
    var btn = el('button', {
        type: 'button',
        style: [
            'display:inline-flex;align-items:center;gap:6px',
            'padding:6px 16px;border:1px solid var(--spButtonBorder,#8a8886)',
            'border-radius:4px;background:var(--spButtonBackground,#fff)',
            'color:var(--spPageText,#323130);font-size:13px;font-family:"Segoe UI",sans-serif',
            'cursor:pointer;transition:background .15s'
        ].join(';')
    }, props.buttonLabel);
    btn.addEventListener('mouseenter', function () { btn.style.background = 'var(--spButtonBackgroundHovered,#f3f2f1)'; });
    btn.addEventListener('mouseleave', function () { btn.style.background = 'var(--spButtonBackground,#fff)'; });
    btn.addEventListener('click', function () { return openPanel(props); });
    container.appendChild(btn);
}
/* ------------------------------------------------------------------ */
/*  Modal panel                                                        */
/* ------------------------------------------------------------------ */
function openPanel(props) {
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var crumbs, overlay, panel, header, closeBtn, breadcrumbBar, content, onKey;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    crumbs = [];
                    overlay = el('div', {
                        style: [
                            'position:fixed;inset:0;z-index:1000001',
                            'background:rgba(0,0,0,.4);display:flex;justify-content:flex-end'
                        ].join(';')
                    });
                    panel = el('div', {
                        style: [
                            'width:420px;max-width:100vw;height:100%;background:var(--spPageBackground,#fff)',
                            'box-shadow:-4px 0 24px rgba(0,0,0,.18);display:flex;flex-direction:column',
                            'font-family:"Segoe UI","Segoe UI Web",sans-serif;color:var(--spPageText,#323130)',
                            'animation:spfxFpSlideIn .2s ease-out'
                        ].join(';')
                    });
                    // Inject keyframe animation
                    injectPanelStyles();
                    header = el('div', {
                        style: 'display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid #edebe9'
                    });
                    header.appendChild(el('span', { style: 'font-size:18px;font-weight:600' }, 'Select a PDF'));
                    closeBtn = el('button', {
                        type: 'button',
                        style: 'background:none;border:none;font-size:18px;cursor:pointer;padding:4px 8px;color:var(--spPageText,#605e5c)',
                        'aria-label': 'Close'
                    }, '✕');
                    closeBtn.addEventListener('click', function () { return overlay.remove(); });
                    header.appendChild(closeBtn);
                    panel.appendChild(header);
                    breadcrumbBar = el('div', {
                        style: 'padding:10px 20px;font-size:12px;color:#605e5c;border-bottom:1px solid #f3f2f1;min-height:18px'
                    });
                    panel.appendChild(breadcrumbBar);
                    content = el('div', {
                        style: 'flex:1;overflow-y:auto;padding:0'
                    });
                    panel.appendChild(content);
                    overlay.appendChild(panel);
                    // Close on overlay background click
                    overlay.addEventListener('click', function (e) {
                        if (e.target === overlay)
                            overlay.remove();
                    });
                    onKey = function (e) {
                        if (e.key === 'Escape') {
                            overlay.remove();
                            document.removeEventListener('keydown', onKey);
                        }
                    };
                    document.addEventListener('keydown', onKey);
                    document.body.appendChild(overlay);
                    // --- Load library list (initial view)
                    return [4 /*yield*/, showLibraries(content, breadcrumbBar, crumbs, props, overlay)];
                case 1:
                    // --- Load library list (initial view)
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/* ------------------------------------------------------------------ */
/*  Library list view                                                  */
/* ------------------------------------------------------------------ */
function showLibraries(content, breadcrumbBar, crumbs, props, overlay) {
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var url, resp, data, _loop_1, _i, _a, lib, err_1;
        return tslib_1.__generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    crumbs.length = 0;
                    renderBreadcrumbs(breadcrumbBar, crumbs, content, props, overlay);
                    showLoading(content);
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 4, , 5]);
                    url = "".concat(props.webAbsoluteUrl, "/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false") +
                        "&$select=Title,Id,RootFolder/ServerRelativeUrl&$expand=RootFolder&$orderby=Title";
                    return [4 /*yield*/, props.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                case 2:
                    resp = _b.sent();
                    return [4 /*yield*/, resp.json()];
                case 3:
                    data = _b.sent();
                    content.innerHTML = '';
                    if (!data.value.length) {
                        content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#605e5c' }, 'No document libraries found.'));
                        return [2 /*return*/];
                    }
                    _loop_1 = function (lib) {
                        var row = makeRow('📁', lib.Title, '');
                        row.addEventListener('click', function () {
                            crumbs.push({ label: lib.Title, serverRelativeUrl: lib.RootFolder.ServerRelativeUrl });
                            showFolder(content, breadcrumbBar, crumbs, lib.RootFolder.ServerRelativeUrl, props, overlay)
                                .catch(function (err) { return console.error('[FilePicker] Error loading folder:', err); });
                        });
                        content.appendChild(row);
                    };
                    for (_i = 0, _a = data.value; _i < _a.length; _i++) {
                        lib = _a[_i];
                        _loop_1(lib);
                    }
                    return [3 /*break*/, 5];
                case 4:
                    err_1 = _b.sent();
                    console.error('[FilePicker] Error loading libraries:', err_1);
                    content.innerHTML = '';
                    content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#a4262c' }, 'Unable to load document libraries.'));
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    });
}
/* ------------------------------------------------------------------ */
/*  Folder contents view                                               */
/* ------------------------------------------------------------------ */
function showFolder(content, breadcrumbBar, crumbs, folderUrl, props, overlay) {
    return tslib_1.__awaiter(this, void 0, void 0, function () {
        var extensions, _a, foldersResp, filesResp, foldersData, filesData, filteredFiles, _loop_2, _i, _b, folder, _loop_3, _c, filteredFiles_1, file, err_2;
        var _d;
        return tslib_1.__generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    renderBreadcrumbs(breadcrumbBar, crumbs, content, props, overlay);
                    showLoading(content);
                    extensions = ((_d = props.accepts) !== null && _d !== void 0 ? _d : ['.pdf']).map(function (e) { return e.toLowerCase(); });
                    _e.label = 1;
                case 1:
                    _e.trys.push([1, 5, , 6]);
                    return [4 /*yield*/, Promise.all([
                            props.spHttpClient.get("".concat(props.webAbsoluteUrl, "/_api/web/GetFolderByServerRelativeUrl('").concat(encodeURIComponent(folderUrl), "')/Folders?$select=Name,ServerRelativeUrl&$orderby=Name&$filter=Name ne 'Forms'"), sp_http_1.SPHttpClient.configurations.v1),
                            props.spHttpClient.get("".concat(props.webAbsoluteUrl, "/_api/web/GetFolderByServerRelativeUrl('").concat(encodeURIComponent(folderUrl), "')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=Name"), sp_http_1.SPHttpClient.configurations.v1)
                        ])];
                case 2:
                    _a = _e.sent(), foldersResp = _a[0], filesResp = _a[1];
                    return [4 /*yield*/, foldersResp.json()];
                case 3:
                    foldersData = _e.sent();
                    return [4 /*yield*/, filesResp.json()];
                case 4:
                    filesData = _e.sent();
                    filteredFiles = filesData.value.filter(function (f) {
                        var name = f.Name.toLowerCase();
                        return extensions.some(function (ext) { return name.slice(-ext.length) === ext; });
                    });
                    content.innerHTML = '';
                    if (!foldersData.value.length && !filteredFiles.length) {
                        content.appendChild(el('div', {
                            style: 'padding:40px 20px;text-align:center;color:#605e5c'
                        }, extensions.length ? "No ".concat(extensions.join(', '), " files in this folder.") : 'This folder is empty.'));
                        return [2 /*return*/];
                    }
                    _loop_2 = function (folder) {
                        var row = makeRow('📁', folder.Name, '');
                        row.addEventListener('click', function () {
                            crumbs.push({ label: folder.Name, serverRelativeUrl: folder.ServerRelativeUrl });
                            showFolder(content, breadcrumbBar, crumbs, folder.ServerRelativeUrl, props, overlay)
                                .catch(function (err) { return console.error('[FilePicker] Error loading folder:', err); });
                        });
                        content.appendChild(row);
                    };
                    // Render sub-folders
                    for (_i = 0, _b = foldersData.value; _i < _b.length; _i++) {
                        folder = _b[_i];
                        _loop_2(folder);
                    }
                    _loop_3 = function (file) {
                        var sizeKb = Math.round(parseInt(file.Length, 10) / 1024);
                        var date = new Date(file.TimeLastModified);
                        var meta = "".concat(formatFileSize(sizeKb), "  \u00B7  ").concat(date.toLocaleDateString());
                        var row = makeRow('📄', file.Name, meta);
                        row.addEventListener('click', function () {
                            var origin = new URL(props.webAbsoluteUrl).origin;
                            var result = {
                                fileAbsoluteUrl: "".concat(origin).concat(file.ServerRelativeUrl),
                                fileName: file.Name,
                                fileNameWithoutExtension: file.Name.replace(/\.[^.]+$/, '')
                            };
                            props.onSelect(result);
                            overlay.remove();
                        });
                        content.appendChild(row);
                    };
                    // Render files
                    for (_c = 0, filteredFiles_1 = filteredFiles; _c < filteredFiles_1.length; _c++) {
                        file = filteredFiles_1[_c];
                        _loop_3(file);
                    }
                    return [3 /*break*/, 6];
                case 5:
                    err_2 = _e.sent();
                    console.error('[FilePicker] Error loading folder contents:', err_2);
                    content.innerHTML = '';
                    content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#a4262c' }, 'Unable to load folder contents.'));
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    });
}
/* ------------------------------------------------------------------ */
/*  Breadcrumbs                                                        */
/* ------------------------------------------------------------------ */
function renderBreadcrumbs(bar, crumbs, content, props, overlay) {
    bar.innerHTML = '';
    // Root link ("Document Libraries")
    var root = el('span', {
        style: "cursor:pointer;color:#0078d4;".concat(crumbs.length ? '' : 'font-weight:600;color:var(--spPageText,#323130);cursor:default;')
    }, 'Document Libraries');
    if (crumbs.length) {
        root.addEventListener('click', function () {
            showLibraries(content, bar, crumbs, props, overlay)
                .catch(function (err) { return console.error('[FilePicker]', err); });
        });
    }
    bar.appendChild(root);
    // Each breadcrumb segment
    crumbs.forEach(function (crumb, i) {
        bar.appendChild(el('span', { style: 'margin:0 6px;color:#a19f9d' }, '›'));
        var isLast = i === crumbs.length - 1;
        var link = el('span', {
            style: "cursor:pointer;color:#0078d4;".concat(isLast ? 'font-weight:600;color:var(--spPageText,#323130);cursor:default;' : '')
        }, crumb.label);
        if (!isLast) {
            link.addEventListener('click', function () {
                crumbs.length = i + 1;
                showFolder(content, bar, crumbs, crumb.serverRelativeUrl, props, overlay)
                    .catch(function (err) { return console.error('[FilePicker]', err); });
            });
        }
        bar.appendChild(link);
    });
}
/* ------------------------------------------------------------------ */
/*  UI helpers                                                         */
/* ------------------------------------------------------------------ */
function makeRow(icon, label, meta) {
    var row = el('div', {
        style: [
            'display:flex;align-items:center;gap:10px;padding:10px 20px',
            'cursor:pointer;border-bottom:1px solid #f3f2f1;transition:background .1s'
        ].join(';')
    });
    row.addEventListener('mouseenter', function () { row.style.background = '#f3f2f1'; });
    row.addEventListener('mouseleave', function () { row.style.background = ''; });
    row.appendChild(el('span', { style: 'font-size:20px;flex-shrink:0;width:24px;text-align:center' }, icon));
    var text = el('div', { style: 'flex:1;min-width:0' });
    text.appendChild(el('div', {
        style: 'font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis',
        title: label
    }, label));
    if (meta) {
        text.appendChild(el('div', { style: 'font-size:11px;color:#a19f9d;margin-top:2px' }, meta));
    }
    row.appendChild(text);
    // Chevron for folders, nothing for files
    if (icon === '📁') {
        row.appendChild(el('span', { style: 'color:#a19f9d;font-size:12px;flex-shrink:0' }, '›'));
    }
    return row;
}
function showLoading(container) {
    container.innerHTML = '';
    var loader = el('div', {
        style: 'display:flex;align-items:center;justify-content:center;padding:60px 20px;color:#605e5c;font-size:13px;gap:8px'
    });
    loader.appendChild(el('span', { style: 'animation:spfxFpSpin 1s linear infinite;display:inline-block;font-size:16px' }, '⟳'));
    loader.appendChild(el('span', {}, 'Loading…'));
    container.appendChild(loader);
}
function formatFileSize(kb) {
    if (kb < 1024)
        return "".concat(kb, " KB");
    return "".concat((kb / 1024).toFixed(1), " MB");
}
/** Inject panel animation keyframes (once). */
var stylesInjected = false;
function injectPanelStyles() {
    if (stylesInjected)
        return;
    stylesInjected = true;
    var style = document.createElement('style');
    style.textContent = "\n    @keyframes spfxFpSlideIn {\n      from { transform: translateX(100%); }\n      to   { transform: translateX(0); }\n    }\n    @keyframes spfxFpSpin {\n      from { transform: rotate(0deg); }\n      to   { transform: rotate(360deg); }\n    }\n  ";
    document.head.appendChild(style);
}
/**
 * Tiny DOM element factory.
 */
function el(tag, attrs, text) {
    var node = document.createElement(tag);
    for (var _i = 0, _a = Object.entries(attrs); _i < _a.length; _i++) {
        var _b = _a[_i], k = _b[0], v = _b[1];
        node.setAttribute(k, v);
    }
    if (text !== undefined)
        node.textContent = text;
    return node;
}
//# sourceMappingURL=PropertyPaneFilePicker.js.map