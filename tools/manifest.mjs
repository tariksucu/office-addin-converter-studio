// The two add-in only (XML) manifests, written from addin.config.json and package.json:
//
// - outlook.xml: a mail add-in (MailApp) with a button on the message read and compose surfaces,
//   Mailbox requirement set 1.8 (attachment contents in both modes), ReadWriteItem;
// - office.xml: a task pane add-in (TaskPaneApp) for Word, Excel and PowerPoint with a button on
//   the Home tab, ReadWriteDocument (the export reads the whole document, inserting writes).
//
// Both point every URL at the base (the hosted task pane); nothing else in the add-ins names an
// address. Element order follows Microsoft's schema (and its own templates, which these follow).

/** XML text and attribute values. */
export function esc(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** `1.2.3` → `1.2.3.0`, the four-part version the manifests take. */
export function manifestVersion(version) {
  if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error(`not a version: ${version}`);
  return `${version}.0`;
}

/** The base as given, with exactly one trailing slash, https (or http://localhost for tests). */
export function normalBase(base) {
  const url = new URL(base);
  const local = url.hostname === 'localhost' || url.hostname === '127.0.0.1';
  if (url.protocol !== 'https:' && !(local && url.protocol === 'http:'))
    throw new Error(`the add-ins must be served over https: ${base}`);
  if (url.search || url.hash) throw new Error(`a base has no query or fragment: ${base}`);
  return url.href.endsWith('/') ? url.href : `${url.href}/`;
}

const ICONS = [16, 32, 80];

function resources(base, strings) {
  const images = ICONS.map(
    (size) =>
      `      <bt:Image id="Icon.${size}" DefaultValue="${esc(`${base}assets/icon-${size}.png`)}"/>`,
  ).join('\n');
  const short = Object.entries(strings.short)
    .map(([id, value]) => `      <bt:String id="${id}" DefaultValue="${esc(value)}"/>`)
    .join('\n');
  const long = Object.entries(strings.long)
    .map(([id, value]) => `      <bt:String id="${id}" DefaultValue="${esc(value)}"/>`)
    .join('\n');
  const urls = Object.entries(strings.urls)
    .map(([id, value]) => `      <bt:Url id="${id}" DefaultValue="${esc(value)}"/>`)
    .join('\n');
  return `    <Resources>
      <bt:Images>
${images}
      </bt:Images>
      <bt:Urls>
${urls}
      </bt:Urls>
      <bt:ShortStrings>
${short}
      </bt:ShortStrings>
      <bt:LongStrings>
${long}
      </bt:LongStrings>
    </Resources>`;
}

function icon(indent) {
  return [
    `${indent}<Icon>`,
    ...ICONS.map((size) => `${indent}  <bt:Image size="${size}" resid="Icon.${size}"/>`),
    `${indent}</Icon>`,
  ].join('\n');
}

function header(config, app, version, iconSizes, base) {
  return `  <Id>${esc(app.id)}</Id>
  <Version>${manifestVersion(version)}</Version>
  <ProviderName>${esc(config.providerName)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${esc(app.name)}"/>
  <Description DefaultValue="${esc(app.description)}"/>
  <IconUrl DefaultValue="${esc(`${base}assets/icon-${iconSizes[0]}.png`)}"/>
  <HighResolutionIconUrl DefaultValue="${esc(`${base}assets/icon-${iconSizes[1]}.png`)}"/>
  <SupportUrl DefaultValue="${esc(config.supportUrl)}"/>`;
}

/** The Outlook add-in's manifest. */
export function outlookManifest(config, version, baseUrl = config.baseUrl) {
  const base = normalBase(baseUrl);
  const app = config.outlook;
  const pane = `${base}taskpane.html`;
  const button = (surface, id, group) => `          <ExtensionPoint xsi:type="${surface}">
            <OfficeTab id="TabDefault">
              <Group id="${group}">
                <Label resid="Group.Label"/>
                <Control xsi:type="Button" id="${id}">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="${id}.Tooltip"/>
                  </Supertip>
${icon('                  ')}
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" xsi:type="MailApp">
${header(config, app, version, [64, 128], base)}
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.8"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${esc(pane)}"/>
        <RequestedHeight>360</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="${esc(pane)}"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>
  <DisableEntityHighlighting>true</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.8">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
${button('MessageReadCommandSurface', 'ReadButton', 'ReadGroup')}
${button('MessageComposeCommandSurface', 'ComposeButton', 'ComposeGroup')}
        </DesktopFormFactor>
      </Host>
    </Hosts>
${resources(base, {
  urls: { 'Commands.Url': `${base}commands.html`, 'Taskpane.Url': pane },
  short: { 'Group.Label': 'Converter Studio', 'Button.Label': 'Convert attachments' },
  long: {
    'ReadButton.Tooltip':
      'Convert this message’s attachments with Converter Studio (HEIC to JPG, Office files to PDF, smaller PDFs and more), then reply with the result or download it.',
    'ComposeButton.Tooltip':
      'Fit the attachments of the message you are writing for sending (photos made smaller on your computer), or convert them with Converter Studio and replace them or attach the result.',
  },
})}
  </VersionOverrides>
</OfficeApp>
`;
}

/** The Word, Excel and PowerPoint add-in's manifest. */
export function officeManifest(config, version, baseUrl = config.baseUrl) {
  const base = normalBase(baseUrl);
  const app = config.office;
  const host = (type) => `      <Host xsi:type="${type}">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="ConverterStudioGroup">
                <Label resid="Group.Label"/>
${icon('                ')}
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
${icon('                  ')}
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ConverterStudioPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
${header(config, app, version, [32, 64], base)}
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${esc(`${base}taskpane.html`)}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
${['Document', 'Workbook', 'Presentation'].map(host).join('\n')}
    </Hosts>
${resources(base, {
  urls: {
    'GetStarted.LearnMoreUrl': config.supportUrl,
    'Commands.Url': `${base}commands.html`,
    'Taskpane.Url': `${base}taskpane.html`,
  },
  short: {
    'GetStarted.Title': 'Converter Studio is ready',
    'Group.Label': 'Converter Studio',
    'Button.Label': 'Convert files',
  },
  long: {
    'GetStarted.Description':
      'On the Home tab, select Convert files to insert a picture Office cannot open, insert the text of a scan, or export this document.',
    'Button.Tooltip':
      'Insert pictures Office cannot open and the text of scans, and export this document to formats Office does not save, with Converter Studio.',
  },
})}
  </VersionOverrides>
</OfficeApp>
`;
}
