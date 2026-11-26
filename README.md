# Converter Studio for Outlook, Word, Excel and PowerPoint

Two Office add-ins that convert files with [Converter Studio](https://converter.studio), an
online file converter, without leaving Office.

- **Attachment Converter for Converter Studio** (Outlook):
  - **Fit attachments for sending**: while you write a message, bring all of its attachments under
    16 MB or 10 MB and replace them in place. Photos are made smaller on your computer, at no
    charge and without an account (see [below](#fit-attachments-for-sending)).
  - Convert the file attachments of a message to any format Converter Studio makes of them: HEIC
    photos to JPG, Word, Excel and PowerPoint files to PDF, a PDF to a smaller PDF, a photo of a
    page to text, and more. While you write a message, the converted file **replaces the
    original** attachment or is **attached beside it**; in a message you received, you **reply
    with the converted file** attached, or download it.
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

## Fit attachments for sending

Open **Convert attachments** in the message you are writing, choose the total and select **Fit
attachments**:

- **16 MB** gets through Gmail and Yahoo Mail (25 MB once a mail program has encoded the
  attachments, which adds about a third); **10 MB** also gets through Outlook.com and iCloud Mail
  (20 MB). The add-in remembers the choice.
- Cheapest first:
  1. If the attachments are under the total already, nothing changes.
  2. Photos (JPEG, PNG, WebP, and iPhone HEIC/HEIF photos) are made smaller **on your computer**,
     largest first, until the message fits: at most 2048 pixels on the longer side, then lower
     quality and 1600 or 1280 pixels only while the message is still over. Photos are turned
     upright from their EXIF orientation; a PNG without transparency may become a JPG, a picture
     with transparency stays PNG, and a moving PNG or WebP is left as it is. This part is free and
     needs no account or key.
  3. What still does not fit and needs Converter Studio's servers (a PDF, a video, a sound file, a
     photo the computer cannot read) is made smaller there only with your API key: the pane shows
     the price first (**Fit for N credits**), and **Cancel** charges nothing (the photos made
     smaller on the computer are kept). Without a key, the pane says which files a key
     would make smaller and links to your account page.
- Each attachment is replaced only once its smaller version has been added to the message; if
  anything fails, the original stays. Names and order are kept; the extension changes only with
  the type (a HEIC photo becomes a JPG). Pictures inside the message text, attached emails and
  OneDrive links are not touched.
- The pane sums it up in one line, for example "4 attachments, 31.2 MB → 14.8 MB. Fits Gmail and
  Yahoo Mail." If the attachments cannot fit, it names what is still too big: share that from
  OneDrive as a link instead.
- **Photos made smaller lose their metadata (EXIF details), the location included.**

## Connect your account

Fitting photos for sending works without an account. Everything else needs one:

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
  waits for you. **Fit attachments for sending** always shows the price of its server part first;
  its photo part costs nothing. Free accounts get credits every day;
  [pricing](https://converter.studio/pricing) lists the paid plans. The largest file depends on
  your plan.
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
- **Fit attachments for sending** makes photos smaller inside the add-in, on your computer: they
  are not sent anywhere. Only with your key, and only after you accept the price, are the files
  the computer cannot make smaller (PDF, video, sound) sent to Converter Studio. Photos made
  smaller lose their metadata (EXIF details), the location included.
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
  to `https://converter.studio/api` (the same origin, without cookies). The HEIC decoder comes
  from the same place (`vendor/`, see [Third-party code](#third-party-code)).

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
| `src/fit/`                  | Fit attachments for sending: the plan (`budget.ts`), the steps (`fit.ts`), the photo encoders    |
| `src/office/`               | The Office.js calls: Outlook attachments and replies, Word/Excel/PowerPoint insertion and export |
| `src/ui/`, `src/strings.ts` | The task pane and its texts                                                                      |
| `test/`                     | Unit tests with a fake Office.js (`fake-office.ts`) and a fake API (`fake-api.ts`)               |
| `dev/`, `e2e/`              | The pane in a real browser against the local stack, with the stub Office host (SUBMISSION §8)    |
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

## Third-party code

- [libheif-js](https://github.com/catdad-experiments/libheif-js) **1.23.5** (LGPL-3.0; libheif
  itself is LGPL-3.0), the WebAssembly build of libheif that decodes HEIC/HEIF photos for **Fit
  attachments for sending**. It is pinned in `package.json` and copied by `tools/build.mjs`, as
  published on npm and unmodified, to `dist/vendor/libheif-js-1.23.5/`: `libheif-wasm/libheif.js`,
  `libheif-wasm/libheif.wasm`, their licences (`LICENSE`, `libheif-wasm/LICENSE`), `README.md` and
  `package.json`. It is a separate file that the pane loads only when a HEIC photo has to be made
  smaller, never bundled into the pane's code, so it can be replaced by any other build of the
  same version. Source: <https://www.npmjs.com/package/libheif-js/v/1.23.5> and
  <https://github.com/strukturag/libheif>.

## Licence

MIT, see [LICENSE](LICENSE). Copyright Converter Studio. Third-party code keeps its own licence
(above).
