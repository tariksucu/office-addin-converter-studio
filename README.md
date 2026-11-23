# Converter Studio for Outlook, Word, Excel and PowerPoint

Two Office add-ins that convert files with [Converter Studio](https://converter.studio), an
online file converter, without leaving Office.

- **Attachment Converter for Converter Studio** (Outlook): convert the file attachments of a
  message to any format Converter Studio makes of them: HEIC photos to JPG, Word, Excel and
  PowerPoint files to PDF, a PDF to a smaller PDF, a photo of a page to text, and more.
  - While you write a message, the converted file **replaces the original** attachment or is
    **attached beside it**.
  - In a message you received, you **reply with the converted file** attached, or download it.
- **File Converter for Converter Studio** (Word, Excel, PowerPoint): things Office does not do
  itself.
  - **Insert a picture** Office cannot open: HEIC (older Office), camera RAW (DNG, CR2, NEF,
    ARW), AVIF, WebP, TIFF, SVG, EPS and Illustrator files, or one page of a PDF. Photos become
    JPG at the size you choose; drawings, PDF pages and pictures that may be see-through become
    PNG. Word inserts it at the cursor, Excel at the selected cell, PowerPoint on the slide.
  - **Text from a scan**: the text of a scanned PDF or a photo of a page, recognised in the
    language found on the page or the one you pick, inserted at the cursor (Word), from the
    selected cell down (Excel) or into the selected text box (PowerPoint).
  - **Export this document** to a format the app does not save itself: Word to EPUB,
    Markdown, HTML, RTF, plain text, DOC or PDF/A; Excel to JSON, XLS, a PDF with each sheet on
    one page, or PDF/A; PowerPoint to PPT, a PDF with speaker notes, a handout PDF or PDF/A.
    The list is whatever Converter Studio's servers offer for the document's format at that
    moment.

## Install

Get the add-in from Microsoft Marketplace (in Outlook: **Apps** or **Get Add-ins**; in Word,
Excel and PowerPoint: **Home** > **Add-ins**), search for Converter Studio and select **Add**.
Microsoft 365 admins can deploy it to their organisation from the Microsoft 365 admin center
(**Settings** > **Integrated apps**).

The add-ins work in:

- Outlook on the web (Microsoft 365 and Outlook.com), new Outlook on Windows, classic Outlook on
  Windows (Microsoft 365, Outlook 2021 and later) and Outlook on Mac (new interface), with
  Exchange Online mailboxes: the add-in needs Mailbox requirement set 1.8, which on-premises
  Exchange servers do not offer. Replying with a converted file needs set 1.15 (Outlook on the
  web, new Outlook, classic Outlook from version 2412); elsewhere, Outlook on Mac included, the
  converted file is downloaded instead.
- Word, Excel and PowerPoint on the web, on Windows (Microsoft 365, Office 2021 and later) and
  on Mac.

## Connect your account

1. [Create a Converter Studio account](https://converter.studio/sign-up) (it starts with daily
   credits at no charge) and create an API key on your
   [account page](https://converter.studio/account#api-keys). Keys start with `cvs_live_`.
2. Open the add-in (**Convert attachments** in Outlook, **Convert files** on the **Home** tab of
   Word, Excel and PowerPoint), paste the key and select **Connect**.

The pane then shows your account's e-mail address and the credits you have left.
**Disconnect** forgets the key on this device.

## Credits and limits

- Each conversion uses credits of your Converter Studio account: most files cost 1 credit.
  Conversions of up to 3 credits start at once; for a dearer one the pane shows the price and
  waits for you. Free accounts get credits every day; [pricing](https://converter.studio/pricing)
  lists the paid plans. The largest file depends on your plan.
- Outlook takes attachments of up to 25 MB from an add-in.
- The export converts the document as it was last saved.
- PowerPoint inserts text into the selected text box or shape: select one first.
- Converted files can be downloaded for 6 hours.

## Privacy and security

- Only what you choose is sent to Converter Studio: the attachment you convert, the picture or
  scan you pick, or, when you export, the document itself. Converter Studio deletes each upload
  once it is converted and each result after 6 hours. See the
  [privacy policy](https://converter.studio/privacy) and the
  [terms](https://converter.studio/terms).
- The add-in sends nothing before you connect, and collects no usage data. Its requests name it
  (`X-Converter-Studio-Client: office/<version>`) so Converter Studio can count conversions by
  add-in.
- The API key is kept in the add-in's own browser storage on this device, under the partition
  key Office gives the add-in, as Microsoft recommends for private data. It is not kept in
  Outlook's roaming settings (Microsoft says not to keep credentials there) or in a document's
  settings (those travel with the file). Whoever holds the key can see your account's e-mail
  address and credits and spend your credits; delete it on your account page if you no longer
  use it.
- The task pane is served from `https://converter.studio/integrations/office/` and talks only
  to `https://converter.studio/api` (the same origin, without cookies).

## Help

[converter.studio/support](https://converter.studio/support) or support@converter.studio.

## Development

A static bundle (TypeScript, no framework) and two add-in only (XML) manifests.

| Path                        | What                                                                                             |
| --------------------------- | ------------------------------------------------------------------------------------------------ |
| `addin.config.json`         | The one place for the add-ins' address (`baseUrl`), ids, names and descriptions                  |
| `manifest/`                 | The manifests for `baseUrl`, generated by `tools/build.mjs` (committed)                          |
| `src/api.ts`                | Converter Studio REST API client (formats, account, uploads, conversions, downloads)             |
| `src/catalog.ts`            | What each file becomes, from the API's format list (`GET /v1/formats`)                           |
| `src/convert.ts`            | One conversion: upload, convert (asking above a price), wait, result                             |
| `src/office/`               | The Office.js calls: Outlook attachments and replies, Word/Excel/PowerPoint insertion and export |
| `src/ui/`, `src/strings.ts` | The task pane and its texts                                                                      |
| `test/`                     | Unit tests with a fake Office.js (`fake-office.ts`) and a fake API (`fake-api.ts`)               |
| `tools/`                    | The build (esbuild), the manifest generator and the icons                                        |

```sh
npm ci
npm test            # unit tests (vitest; the pane's tests run in jsdom)
npm run typecheck
npm run lint        # Microsoft's eslint-plugin-office-addins rules and typescript-eslint
npm run build       # dist/: the pane for baseUrl, and dist/manifest/
npm run validate    # Microsoft's manifest validator (office-addin-manifest validate -p)
node tools/build.mjs --base https://example.com/addin/ --out /tmp/addin   # another address
```

`npm run validate` sends the manifests to Microsoft's validation service, which also fetches
the icons from `baseUrl`: it passes once the bundle is served there.

## Licence

MIT, see [LICENSE](LICENSE). Copyright Converter Studio.
