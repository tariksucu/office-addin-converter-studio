# Changelog

## 1.1.0 - unreleased

- Outlook: **Fit attachments for sending**, while you write a message. One button brings all of
  the draft's attachments under a total, 16 MB (gets through Gmail and Yahoo Mail; the same size
  as converter.studio's "Fit for email") or 10 MB (also Outlook.com and iCloud Mail), and replaces
  them in place. The choice is remembered.
  - Cheapest first: nothing changes when the draft fits already; photos (JPEG, PNG, WebP, and
    HEIC/HEIF through libheif-js) are made smaller on the computer first, largest first, at most
    2048 px at quality 0.85, then by quality and at 1600 or 1280 px only while the draft is still
    over. EXIF orientation is applied, an opaque PNG may become a JPG, a see-through picture stays
    PNG, a moving PNG or WebP is left as it is. This part is free and works without an account or
    key.
  - What still does not fit (PDF, video, sound, a photo the computer cannot read) goes to
    Converter Studio only with an API key, after the price (the API's own quote) is shown: "Fit for N
    credits" goes ahead, "Cancel" charges nothing and keeps the photos made smaller on the
    computer (the same labels as the Thunderbird add-on). Without a key, the pane says what a key would do and links to the account
    page.
  - A video becomes a smaller MP4 and a sound a smaller MP3; an MP4 stays an MP4 and an MP3 an
    MP3, with the same name (Converter Studio makes them smaller in their own format, the
    `compress` of its format list). Only from a server that does not offer that does an MP4
    become a MOV and an MP3 an M4A, as before.
  - A file the servers cannot make smaller comes back as it was and costs nothing: the pane keeps
    the original and says so in the outcome line ("Already as small as it gets, so kept as it was
    (nothing charged): …"); it is not reported as a failure.
  - Each original is replaced only once its result has been added; a failure leaves the original.
    Names and order are kept; the extension changes only with the type (HEIC → JPG). Pictures in
    the body, attached emails and OneDrive links are not touched.
  - The outcome is one line ("4 attachments, 31.2 MB → 14.8 MB. Fits Gmail and Yahoo Mail."); a
    draft that cannot fit names what is still too big and points to sharing it from OneDrive.
  - Re-encoded photos lose their metadata, the location included (said in the pane).
- The HEIC decoder, libheif-js 1.23.5 (LGPL-3.0), is served unmodified next to the pane
  (`vendor/libheif-js-1.23.5/`) and loaded only when a HEIC photo has to be made smaller.
- Requires nothing new from Outlook: Mailbox 1.8, as before. The rest of the add-in (converting
  one attachment, Word, Excel and PowerPoint) is unchanged.

## 1.0.0 - unreleased

- First release: two Office add-ins, served from https://converter.studio/integrations/office/.
- Outlook (Attachment Converter for Converter Studio): convert a message's file attachments to
  any format Converter Studio offers for them; replace the original or attach the result while
  writing; reply with the result (Mailbox 1.15) or download it when reading.
- Word, Excel and PowerPoint (File Converter for Converter Studio): insert pictures Office
  cannot open (HEIC, camera RAW, AVIF, WebP, TIFF, SVG, EPS, Illustrator, a PDF page), insert
  the text of a scan (19 languages or detected), export the document to formats the app does
  not save (EPUB, Markdown, HTML, RTF, TXT, DOC, JSON, XLS, PPT, PDF/A, handouts, notes pages).
- The API key in the add-in's partitioned browser storage; Disconnect forgets it.
