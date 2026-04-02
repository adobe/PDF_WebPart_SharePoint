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

import {
  type IPropertyPaneField,
  type IPropertyPaneCustomFieldProps,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

/* ------------------------------------------------------------------ */
/*  Public interfaces                                                  */
/* ------------------------------------------------------------------ */

export interface IFilePickerResult {
  fileAbsoluteUrl: string;
  fileName: string;
  fileNameWithoutExtension: string;
}

export interface IPropertyPaneFilePickerProps {
  /** Unique key for the field */
  key: string;
  /** Label shown above the button */
  label: string;
  /** Text on the browse button */
  buttonLabel: string;
  /** File extensions to show (e.g. ['.pdf']) */
  accepts?: string[];
  /** Currently selected file */
  value: IFilePickerResult | undefined;
  /** Absolute URL of the current site */
  webAbsoluteUrl: string;
  /** SPHttpClient from the web part context */
  spHttpClient: SPHttpClient;
  /** Callback when a file is selected */
  onSelect: (result: IFilePickerResult) => void;
}

/* ------------------------------------------------------------------ */
/*  REST response shapes                                               */
/* ------------------------------------------------------------------ */

interface ILibraryInfo {
  Title: string;
  Id: string;
  RootFolder: { ServerRelativeUrl: string };
}

interface IFolderInfo {
  Name: string;
  ServerRelativeUrl: string;
}

interface IFileInfo {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified: string;
  Length: string;
}

interface IBreadcrumb {
  label: string;
  serverRelativeUrl: string;
}

/* ------------------------------------------------------------------ */
/*  Factory function                                                   */
/* ------------------------------------------------------------------ */

export function PropertyPaneFilePicker(
  targetProperty: string,
  props: IPropertyPaneFilePickerProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: props.key,
      onRender: (domElement: HTMLElement) => {
        renderField(domElement, props);
      },
      onDispose: (domElement: HTMLElement) => {
        domElement.innerHTML = '';
      }
    }
  };
}

/* ------------------------------------------------------------------ */
/*  Property-pane field renderer                                       */
/* ------------------------------------------------------------------ */

function renderField(
  container: HTMLElement,
  props: IPropertyPaneFilePickerProps
): void {
  container.innerHTML = '';

  // -- Label
  const label = el('label', {
    style: 'display:block;font-weight:600;font-size:14px;padding-bottom:5px;color:var(--spPageText,#323130);font-family:"Segoe UI",sans-serif'
  }, props.label);
  container.appendChild(label);

  // -- Selected file display
  const fileName = props.value?.fileName ?? '';
  if (fileName) {
    const selected = el('div', {
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
  const btn = el('button', {
    type: 'button',
    style: [
      'display:inline-flex;align-items:center;gap:6px',
      'padding:6px 16px;border:1px solid var(--spButtonBorder,#8a8886)',
      'border-radius:4px;background:var(--spButtonBackground,#fff)',
      'color:var(--spPageText,#323130);font-size:13px;font-family:"Segoe UI",sans-serif',
      'cursor:pointer;transition:background .15s'
    ].join(';')
  }, props.buttonLabel);

  btn.addEventListener('mouseenter', () => { btn.style.background = 'var(--spButtonBackgroundHovered,#f3f2f1)'; });
  btn.addEventListener('mouseleave', () => { btn.style.background = 'var(--spButtonBackground,#fff)'; });
  btn.addEventListener('click', () => openPanel(props));

  container.appendChild(btn);
}

/* ------------------------------------------------------------------ */
/*  Modal panel                                                        */
/* ------------------------------------------------------------------ */

async function openPanel(props: IPropertyPaneFilePickerProps): Promise<void> {
  // Breadcrumb state
  const crumbs: IBreadcrumb[] = [];

  // --- Overlay
  const overlay = el('div', {
    style: [
      'position:fixed;inset:0;z-index:1000001',
      'background:rgba(0,0,0,.4);display:flex;justify-content:flex-end'
    ].join(';')
  });

  // --- Panel container
  const panel = el('div', {
    style: [
      'width:420px;max-width:100vw;height:100%;background:var(--spPageBackground,#fff)',
      'box-shadow:-4px 0 24px rgba(0,0,0,.18);display:flex;flex-direction:column',
      'font-family:"Segoe UI","Segoe UI Web",sans-serif;color:var(--spPageText,#323130)',
      'animation:spfxFpSlideIn .2s ease-out'
    ].join(';')
  });

  // Inject keyframe animation
  injectPanelStyles();

  // --- Header
  const header = el('div', {
    style: 'display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid #edebe9'
  });
  header.appendChild(el('span', { style: 'font-size:18px;font-weight:600' }, 'Select a PDF'));
  const closeBtn = el('button', {
    type: 'button',
    style: 'background:none;border:none;font-size:18px;cursor:pointer;padding:4px 8px;color:var(--spPageText,#605e5c)',
    'aria-label': 'Close'
  }, '✕');
  closeBtn.addEventListener('click', () => overlay.remove());
  header.appendChild(closeBtn);
  panel.appendChild(header);

  // --- Breadcrumb bar
  const breadcrumbBar = el('div', {
    style: 'padding:10px 20px;font-size:12px;color:#605e5c;border-bottom:1px solid #f3f2f1;min-height:18px'
  });
  panel.appendChild(breadcrumbBar);

  // --- Content area
  const content = el('div', {
    style: 'flex:1;overflow-y:auto;padding:0'
  });
  panel.appendChild(content);

  overlay.appendChild(panel);

  // Close on overlay background click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) overlay.remove();
  });

  // Close on Escape
  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      overlay.remove();
      document.removeEventListener('keydown', onKey);
    }
  };
  document.addEventListener('keydown', onKey);

  document.body.appendChild(overlay);

  // --- Load library list (initial view)
  await showLibraries(content, breadcrumbBar, crumbs, props, overlay);
}

/* ------------------------------------------------------------------ */
/*  Library list view                                                  */
/* ------------------------------------------------------------------ */

async function showLibraries(
  content: HTMLElement,
  breadcrumbBar: HTMLElement,
  crumbs: IBreadcrumb[],
  props: IPropertyPaneFilePickerProps,
  overlay: HTMLElement
): Promise<void> {
  crumbs.length = 0;
  renderBreadcrumbs(breadcrumbBar, crumbs, content, props, overlay);
  showLoading(content);

  try {
    const url =
      `${props.webAbsoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false` +
      `&$select=Title,Id,RootFolder/ServerRelativeUrl&$expand=RootFolder&$orderby=Title`;

    const resp: SPHttpClientResponse = await props.spHttpClient.get(
      url, SPHttpClient.configurations.v1
    );
    const data = await resp.json() as { value: ILibraryInfo[] };

    content.innerHTML = '';

    if (!data.value.length) {
      content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#605e5c' }, 'No document libraries found.'));
      return;
    }

    for (const lib of data.value) {
      const row = makeRow('📁', lib.Title, '');
      row.addEventListener('click', () => {
        crumbs.push({ label: lib.Title, serverRelativeUrl: lib.RootFolder.ServerRelativeUrl });
        showFolder(content, breadcrumbBar, crumbs, lib.RootFolder.ServerRelativeUrl, props, overlay)
          .catch(err => console.error('[FilePicker] Error loading folder:', err));
      });
      content.appendChild(row);
    }
  } catch (err) {
    console.error('[FilePicker] Error loading libraries:', err);
    content.innerHTML = '';
    content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#a4262c' }, 'Unable to load document libraries.'));
  }
}

/* ------------------------------------------------------------------ */
/*  Folder contents view                                               */
/* ------------------------------------------------------------------ */

async function showFolder(
  content: HTMLElement,
  breadcrumbBar: HTMLElement,
  crumbs: IBreadcrumb[],
  folderUrl: string,
  props: IPropertyPaneFilePickerProps,
  overlay: HTMLElement
): Promise<void> {
  renderBreadcrumbs(breadcrumbBar, crumbs, content, props, overlay);
  showLoading(content);

  const extensions = (props.accepts ?? ['.pdf']).map(e => e.toLowerCase());

  try {
    // Fetch folders and files in parallel
    const [foldersResp, filesResp] = await Promise.all([
      props.spHttpClient.get(
        `${props.webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Folders?$select=Name,ServerRelativeUrl&$orderby=Name&$filter=Name ne 'Forms'`,
        SPHttpClient.configurations.v1
      ),
      props.spHttpClient.get(
        `${props.webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=Name`,
        SPHttpClient.configurations.v1
      )
    ]);

    const foldersData = await foldersResp.json() as { value: IFolderInfo[] };
    const filesData = await filesResp.json() as { value: IFileInfo[] };

    // Filter files by accepted extensions
    const filteredFiles = filesData.value.filter(f => {
      const name = f.Name.toLowerCase();
      return extensions.some(ext => name.slice(-ext.length) === ext);
    });

    content.innerHTML = '';

    if (!foldersData.value.length && !filteredFiles.length) {
      content.appendChild(el('div', {
        style: 'padding:40px 20px;text-align:center;color:#605e5c'
      }, extensions.length ? `No ${extensions.join(', ')} files in this folder.` : 'This folder is empty.'));
      return;
    }

    // Render sub-folders
    for (const folder of foldersData.value) {
      const row = makeRow('📁', folder.Name, '');
      row.addEventListener('click', () => {
        crumbs.push({ label: folder.Name, serverRelativeUrl: folder.ServerRelativeUrl });
        showFolder(content, breadcrumbBar, crumbs, folder.ServerRelativeUrl, props, overlay)
          .catch(err => console.error('[FilePicker] Error loading folder:', err));
      });
      content.appendChild(row);
    }

    // Render files
    for (const file of filteredFiles) {
      const sizeKb = Math.round(parseInt(file.Length, 10) / 1024);
      const date = new Date(file.TimeLastModified);
      const meta = `${formatFileSize(sizeKb)}  ·  ${date.toLocaleDateString()}`;

      const row = makeRow('📄', file.Name, meta);
      row.addEventListener('click', () => {
        const origin = new URL(props.webAbsoluteUrl).origin;
        const result: IFilePickerResult = {
          fileAbsoluteUrl: `${origin}${file.ServerRelativeUrl}`,
          fileName: file.Name,
          fileNameWithoutExtension: file.Name.replace(/\.[^.]+$/, '')
        };
        props.onSelect(result);
        overlay.remove();
      });
      content.appendChild(row);
    }
  } catch (err) {
    console.error('[FilePicker] Error loading folder contents:', err);
    content.innerHTML = '';
    content.appendChild(el('div', { style: 'padding:40px 20px;text-align:center;color:#a4262c' }, 'Unable to load folder contents.'));
  }
}

/* ------------------------------------------------------------------ */
/*  Breadcrumbs                                                        */
/* ------------------------------------------------------------------ */

function renderBreadcrumbs(
  bar: HTMLElement,
  crumbs: IBreadcrumb[],
  content: HTMLElement,
  props: IPropertyPaneFilePickerProps,
  overlay: HTMLElement
): void {
  bar.innerHTML = '';

  // Root link ("Document Libraries")
  const root = el('span', {
    style: `cursor:pointer;color:#0078d4;${crumbs.length ? '' : 'font-weight:600;color:var(--spPageText,#323130);cursor:default;'}`
  }, 'Document Libraries');
  if (crumbs.length) {
    root.addEventListener('click', () => {
      showLibraries(content, bar, crumbs, props, overlay)
        .catch(err => console.error('[FilePicker]', err));
    });
  }
  bar.appendChild(root);

  // Each breadcrumb segment
  crumbs.forEach((crumb, i) => {
    bar.appendChild(el('span', { style: 'margin:0 6px;color:#a19f9d' }, '›'));
    const isLast = i === crumbs.length - 1;
    const link = el('span', {
      style: `cursor:pointer;color:#0078d4;${isLast ? 'font-weight:600;color:var(--spPageText,#323130);cursor:default;' : ''}`
    }, crumb.label);
    if (!isLast) {
      link.addEventListener('click', () => {
        crumbs.length = i + 1;
        showFolder(content, bar, crumbs, crumb.serverRelativeUrl, props, overlay)
          .catch(err => console.error('[FilePicker]', err));
      });
    }
    bar.appendChild(link);
  });
}

/* ------------------------------------------------------------------ */
/*  UI helpers                                                         */
/* ------------------------------------------------------------------ */

function makeRow(icon: string, label: string, meta: string): HTMLElement {
  const row = el('div', {
    style: [
      'display:flex;align-items:center;gap:10px;padding:10px 20px',
      'cursor:pointer;border-bottom:1px solid #f3f2f1;transition:background .1s'
    ].join(';')
  });
  row.addEventListener('mouseenter', () => { row.style.background = '#f3f2f1'; });
  row.addEventListener('mouseleave', () => { row.style.background = ''; });

  row.appendChild(el('span', { style: 'font-size:20px;flex-shrink:0;width:24px;text-align:center' }, icon));

  const text = el('div', { style: 'flex:1;min-width:0' });
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

function showLoading(container: HTMLElement): void {
  container.innerHTML = '';
  const loader = el('div', {
    style: 'display:flex;align-items:center;justify-content:center;padding:60px 20px;color:#605e5c;font-size:13px;gap:8px'
  });
  loader.appendChild(el('span', { style: 'animation:spfxFpSpin 1s linear infinite;display:inline-block;font-size:16px' }, '⟳'));
  loader.appendChild(el('span', {}, 'Loading…'));
  container.appendChild(loader);
}

function formatFileSize(kb: number): string {
  if (kb < 1024) return `${kb} KB`;
  return `${(kb / 1024).toFixed(1)} MB`;
}

/** Inject panel animation keyframes (once). */
let stylesInjected = false;
function injectPanelStyles(): void {
  if (stylesInjected) return;
  stylesInjected = true;
  const style = document.createElement('style');
  style.textContent = `
    @keyframes spfxFpSlideIn {
      from { transform: translateX(100%); }
      to   { transform: translateX(0); }
    }
    @keyframes spfxFpSpin {
      from { transform: rotate(0deg); }
      to   { transform: rotate(360deg); }
    }
  `;
  document.head.appendChild(style);
}

/**
 * Tiny DOM element factory.
 */
function el(tag: string, attrs: Record<string, string>, text?: string): HTMLElement {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    node.setAttribute(k, v);
  }
  if (text !== undefined) node.textContent = text;
  return node;
}
