/**
 * Every text the pane shows. English is the add-ins' only language for now (the manifests'
 * DefaultLocale, en-US, and the Marketplace listing); another language is one more object of the
 * same shape here, with the same language in the manifests' Override elements and a listing in
 * that language (Microsoft's policy 1100.7: the experience must match in each language).
 */

const en = {
  appName: 'Converter Studio',
  loading: 'Loading…',
  // First run and the key.
  introMail:
    'Convert the attachments of this message: HEIC photos to JPG, Word, Excel and PowerPoint files to PDF, a PDF to a smaller PDF, and many more formats. Replace an attachment while you write, or reply with the converted file.',
  introDocument:
    'Insert pictures Office cannot open (HEIC, camera RAW, AVIF, WebP, EPS, Illustrator, a page of a PDF), insert the text of a scanned page, and export this document to formats Office does not save.',
  introAccount:
    'Converter Studio converts the files on its servers, on your account’s credits. You need an account and an API key:',
  stepSignUp: 'Create an account',
  stepSignUpTail: ' (it starts with daily credits at no charge).',
  stepKey: 'Create an API key',
  stepKeyTail: ' on your account page and paste it below.',
  keyLabel: 'API key',
  keyHint: 'Starts with cvs_live_. It stays on this device until you disconnect.',
  connect: 'Connect',
  connecting: 'Checking the key…',
  badKey: 'That is not a Converter Studio API key. Copy it again from your account page.',
  privacyNote: 'Only the files you choose are sent to Converter Studio.',
  privacyLink: 'Privacy policy',
  help: 'Help',
  // The account.
  account: (email: string, credits: number) =>
    `${email} · ${credits} ${credits === 1 ? 'credit' : 'credits'} left`,
  disconnect: 'Disconnect',
  pricing: 'Plans and credits',
  // Outlook.
  attachments: 'Attachments',
  noAttachments: 'This message has no file attachments.',
  refresh: 'Refresh',
  convertTo: (name: string) => `Convert ${name} to`,
  cannotConvert: 'Converter Studio cannot convert this file.',
  convert: 'Convert',
  replaceOriginal: 'Replace the original',
  addAttachment: 'Attach it',
  replyWith: 'Reply with it',
  download: 'Download',
  attached: (name: string) => `${name} is attached.`,
  replaced: (name: string) => `${name} replaced the original.`,
  replyOpened: 'A reply with the file is open.',
  // Fit attachments for sending (Outlook, while writing).
  fitTitle: 'Fit attachments for sending',
  fitTarget: 'Total size',
  fitTargets: {
    16: '16 MB (gets through Gmail and Yahoo Mail)',
    10: '10 MB (also Outlook.com and iCloud Mail)',
  } as Readonly<Record<number, string>>,
  fitButton: 'Fit attachments',
  fitHint:
    'Makes the attachments of this message smaller until they fit, and replaces them in place. Photos are made smaller on this computer, at no charge and without an account; they lose their metadata, location included.',
  fitHintServer:
    'PDFs, videos and sound that still do not fit are made smaller on Converter Studio’s servers, on your credits, after you see the price.',
  fitHintNoKey:
    'PDFs, videos and sound need Converter Studio’s servers and an API key (see below).',
  fitChecking: 'Checking the attachments…',
  fitPhotos: (done: number, total: number) =>
    `Making photos smaller on this computer… ${done} of ${total}`,
  fitSending: (name: string, percent: number) => `Sending ${name}… ${percent}%`,
  fitPricing: (name: string) => `Getting the price for ${name}…`,
  fitConverting: (name: string) => `Making ${name} smaller…`,
  fitReplacing: 'Replacing the attachments…',
  fitPrice: (credits: number, names: string) =>
    `Making ${names} smaller on Converter Studio’s servers costs ${credits} ${credits === 1 ? 'credit' : 'credits'}.`,
  /** The price question’s answers, as in the Thunderbird add-on: “Fit for 1 credit” / Cancel. */
  fitPay: (credits: number) => `Fit for ${credits} ${credits === 1 ? 'credit' : 'credits'}`,
  fitEmpty: 'This message has no attachments.',
  fitAlready: (target: number, size: string) =>
    `The attachments are already under ${target} MB (${size}). Nothing was changed.`,
  fitSummary: (count: number, before: string, after: string) =>
    `${count} ${count === 1 ? 'attachment' : 'attachments'}, ${before === after ? before : `${before} → ${after}`}.`,
  fitFits: (allFour: boolean) =>
    allFour ? 'Fits Gmail, Yahoo Mail, Outlook.com and iCloud Mail.' : 'Fits Gmail and Yahoo Mail.',
  fitStillOver: (target: number, list: string) => `Still over ${target} MB: ${list}.`,
  fitOneDrive: (count: number) =>
    `Share ${count === 1 ? 'it' : 'them'} as a OneDrive link instead: upload to OneDrive, then attach from there as a link.`,
  fitNeedKey: (names: string) =>
    `With a Converter Studio API key, ${names} could be made smaller on its servers.`,
  fitKeyLink: 'Create an API key',
  fitSkipped: (names: string) => `Not converted, so nothing was charged: ${names}.`,
  /** The servers could not make these smaller (`notSmaller`): as in the Thunderbird add-on. */
  fitKept: (names: string, count: number) =>
    count === 1
      ? `Already as small as it gets, so kept as it was (nothing charged): ${names}.`
      : `Already as small as they get, so kept as they were (nothing charged): ${names}.`,
  fitFailed: (name: string, message: string) => `${name} stays as it was: ${message}`,
  fitRemoved: 'it was removed from the message meanwhile.',
  fitNotSmaller: 'the result was not smaller.',
  /** "a", "a and b", "a, b and c". */
  list: (items: readonly string[]) =>
    items.length <= 1
      ? (items[0] ?? '')
      : `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`,
  // Word, Excel, PowerPoint.
  pictureTitle: 'Insert a picture',
  pictureHint:
    'HEIC, camera RAW (DNG, CR2, NEF, ARW), AVIF, WebP, TIFF, SVG, EPS and Illustrator files, or one page of a PDF.',
  pictureFile: 'Picture or PDF',
  pdfPage: 'PDF page',
  size: 'Size',
  sizeOriginal: 'Original size',
  sizePixels: (pixels: string) => `${pixels} pixels on the longer side`,
  insertPicture: 'Insert picture',
  pictureInserted: 'The picture is inserted.',
  textTitle: 'Text from a scan',
  textHint: {
    word: 'The text of a scanned PDF or a photo of a page goes in at the cursor.',
    excel: 'The text of a scanned PDF or a photo of a page goes in from the selected cell down.',
    powerpoint:
      'The text of a scanned PDF or a photo of a page goes into the selected text box or shape.',
  },
  textFile: 'Scan (PDF or picture)',
  language: 'Language',
  pages: 'Pages (for example 1-3, 5)',
  pagesAll: 'all',
  insertText: 'Insert text',
  textInserted: 'The text is inserted.',
  noText: 'No text was found in it.',
  notOneFile: 'The result is several files in a ZIP, which cannot be inserted. Try one page.',
  exportTitle: 'Export this document',
  exportHint:
    'Converts the saved state of this document into a format Office does not save itself.',
  exportTo: 'Format',
  exportButton: 'Export',
  exportReady: (name: string) => `${name} is ready.`,
  exportNone: 'Converter Studio has nothing to offer for this document right now.',
  chooseFile: 'Choose a file first.',
  notPicture: 'Choose a picture or a PDF from the list of formats above.',
  notScan: 'Choose a scanned PDF or a picture (JPG, PNG, HEIC, TIFF, WebP or BMP).',
  badPage: 'Write the number of the PDF page.',
  badPages: 'Write pages like 1-3, 5, or leave it empty for all pages.',
  // Progress and prices.
  uploading: (percent: number) => `Sending the file… ${percent}%`,
  converting: 'Converting…',
  convertingPercent: (percent: number) => `Converting… ${percent}%`,
  fetching: 'Getting the result…',
  cancel: 'Cancel',
  canceled: 'Canceled.',
  failed: 'Something went wrong. Try again.',
  price: (credits: number) => `This conversion costs ${credits} credits. Convert it?`,
  pay: 'Convert',
  credits: (credits: number) => `${credits} ${credits === 1 ? 'credit' : 'credits'} used.`,
  expires: 'The download link works for 6 hours.',
  tooLargeForOutlook: 'The result is larger than the 25 MB Outlook takes from an add-in.',
  outsideOffice:
    'Open this page from Outlook, Word, Excel or PowerPoint: it is the Converter Studio add-in.',
  unsupportedHost: 'This add-in works in Outlook, Word, Excel and PowerPoint.',
  noMessage: 'Open a message to convert its attachments.',
  tryAgain: 'Try again',
  formatsFailed: 'Converter Studio could not be reached. Check your connection and try again.',
};

export type Messages = typeof en;

const LANGUAGES: Readonly<Record<string, Messages>> = { en };

/** The messages for Office's display language (`en-US`), English when there are none. */
export function messages(language: string): Messages {
  const code = language.toLowerCase().split('-')[0] ?? 'en';
  return LANGUAGES[code] ?? en;
}
