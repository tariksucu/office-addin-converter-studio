import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { JSDOM } from 'jsdom';
import { describe, expect, it } from 'vitest';
import { normalBase, officeManifest, outlookManifest } from '../tools/manifest.mjs';
import config from '../addin.config.json';
import { version } from '../package.json';

const root = join(import.meta.dirname, '..');
const build = { outlook: outlookManifest, office: officeManifest };

function parse(xml: string): Document {
  const parsed = new new JSDOM('').window.DOMParser().parseFromString(xml, 'application/xml');
  expect(parsed.getElementsByTagName('parsererror')).toHaveLength(0);
  return parsed;
}

/** The attribute of every such element that has it. */
const attr = (doc: Document, tag: string, name = 'DefaultValue') =>
  [...doc.getElementsByTagName(tag)]
    .map((e) => e.getAttribute(name))
    .filter((v): v is string => v !== null);

describe('manifests', () => {
  for (const kind of ['outlook', 'office'] as const) {
    describe(kind, () => {
      const xml = build[kind](config, version);
      const doc = parse(xml);

      it('is the committed manifest/ file (npm run manifests writes it)', () => {
        expect(readFileSync(join(root, 'manifest', `${kind}.xml`), 'utf8')).toBe(xml);
      });

      it('carries the configured id, the package version and Converter Studio as provider', () => {
        expect(doc.getElementsByTagName('Id')[0]!.textContent).toBe(config[kind].id);
        expect(doc.getElementsByTagName('Version')[0]!.textContent).toBe(`${version}.0`);
        expect(doc.getElementsByTagName('ProviderName')[0]!.textContent).toBe('Converter Studio');
        expect(attr(doc, 'DisplayName')).toEqual([config[kind].name]);
        expect(attr(doc, 'SupportUrl')).toEqual(['https://converter.studio/support']);
      });

      it('points every file at the hosted base, over https', () => {
        const urls = [
          ...attr(doc, 'IconUrl'),
          ...attr(doc, 'HighResolutionIconUrl'),
          ...attr(doc, 'SourceLocation'),
          ...attr(doc, 'bt:Image'),
          ...attr(doc, 'bt:Url').filter((u) => !u.endsWith('/support')),
        ];
        expect(urls.length).toBeGreaterThan(5);
        for (const url of urls)
          expect(url).toMatch(/^https:\/\/converter\.studio\/integrations\/office\//);
      });

      it('keeps strings within the schema’s lengths', () => {
        expect(config[kind].name.length).toBeLessThanOrEqual(50);
        expect(config[kind].description.length).toBeLessThanOrEqual(250);
        for (const s of doc.getElementsByTagName('bt:ShortStrings')[0]!.children)
          expect(s.getAttribute('DefaultValue')!.length).toBeLessThanOrEqual(125);
        for (const s of doc.getElementsByTagName('bt:LongStrings')[0]!.children)
          expect(s.getAttribute('DefaultValue')!.length).toBeLessThanOrEqual(250);
      });

      it('uses only icons that exist in the right sizes', () => {
        for (const url of [
          ...attr(doc, 'IconUrl'),
          ...attr(doc, 'HighResolutionIconUrl'),
          ...attr(doc, 'bt:Image'),
        ]) {
          const file = url.split('/').pop()!;
          const png = readFileSync(join(root, 'assets', file));
          const size = Number(/icon-(\d+)\.png/.exec(file)![1]);
          expect(png.readUInt32BE(16)).toBe(size);
          expect(png.readUInt32BE(20)).toBe(size);
        }
      });
    });
  }

  it('Outlook: Mailbox 1.8, read and compose buttons, read/write item permission', () => {
    const doc = parse(build.outlook(config, version));
    expect(attr(doc, 'Set', 'MinVersion')).toEqual(['1.8']);
    expect(attr(doc, 'ExtensionPoint', 'xsi:type')).toEqual([
      'MessageReadCommandSurface',
      'MessageComposeCommandSurface',
    ]);
    expect(doc.getElementsByTagName('Permissions')[0]!.textContent).toBe('ReadWriteItem');
    // Policy 1120.4: no ItemSend event and no CustomPane in add-in only manifests.
    expect(build.outlook(config, version)).not.toMatch(/ItemSend|CustomPane|MobileFormFactor/);
  });

  it('Office: Word, Excel and PowerPoint, a Home tab button each, read/write document', () => {
    const doc = parse(build.office(config, version));
    expect(attr(doc, 'Host', 'Name')).toEqual(['Document', 'Workbook', 'Presentation']);
    expect(attr(doc, 'ExtensionPoint', 'xsi:type')).toEqual([
      'PrimaryCommandSurface',
      'PrimaryCommandSurface',
      'PrimaryCommandSurface',
    ]);
    expect(doc.getElementsByTagName('Permissions')[0]!.textContent).toBe('ReadWriteDocument');
  });

  it('builds for another base (a local test server) and refuses plain http elsewhere', () => {
    const local = build.office(config, version, 'http://localhost:8080/integrations/office');
    expect(local).toContain('http://localhost:8080/integrations/office/taskpane.html');
    expect(() => normalBase('http://converter.studio/integrations/office/')).toThrow(/https/);
    expect(() => normalBase('https://converter.studio/x/?a=1')).toThrow(/query/);
  });
});
