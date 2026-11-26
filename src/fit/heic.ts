/**
 * HEIC and HEIF photos (iPhone's), which most webviews cannot read, decoded on the computer with
 * libheif-js 1.23.5 (https://github.com/catdad-experiments/libheif-js, LGPL-3.0, the version
 * apps/extension ships): its WebAssembly build, copied as published on npm
 * (`libheif-wasm/libheif.js`, `libheif.wasm` and their licences) into the pane's `vendor/` folder
 * by tools/build.mjs. It is a separate file, loaded only when a HEIC photo has to be made
 * smaller (about 1.5 MB, once), and never bundled or changed.
 */

/** Where the copy lives, next to taskpane.html. */
export const LIBHEIF_FOLDER = 'vendor/libheif-js-1.23.5/libheif-wasm/';

export interface HeifImage {
  get_width(): number;
  get_height(): number;
  display(
    target: { data: Uint8ClampedArray; width: number; height: number },
    done: (shown: { data: Uint8ClampedArray } | null) => void,
  ): void;
  free(): void;
}

export interface Libheif {
  HeifDecoder: new () => { decode(bytes: Uint8Array): HeifImage[] };
}

/** The factory `libheif.js` defines (Emscripten's module options). */
export type LibheifFactory = (options: {
  locateFile?: (path: string) => string;
  wasmBinary?: Uint8Array;
  onRuntimeInitialized?: () => void;
  onAbort?: (reason: unknown) => void;
  print?: (text: string) => void;
  printErr?: (text: string) => void;
}) => Libheif;

/**
 * Starts the WebAssembly module, its `.wasm` taken from `folder`. This build compiles the module
 * synchronously (Emscripten's WASM_ASYNC_COMPILATION off), so a browser hands it the bytes
 * (`wasmBinary`, fetched first); Node reads them from `folder` itself.
 */
export function startLibheif(
  factory: LibheifFactory,
  folder: string,
  wasmBinary?: Uint8Array,
): Promise<Libheif> {
  return new Promise((resolve, reject) => {
    let lib: Libheif | null = null;
    let ready = false;
    const settle = () => {
      if (ready && lib) resolve(lib);
    };
    lib = factory({
      locateFile: (path) => folder + path,
      ...(wasmBinary ? { wasmBinary } : {}),
      onRuntimeInitialized: () => {
        ready = true;
        settle();
      },
      onAbort: (reason) => reject(new Error(`HEIC decoder failed to start: ${String(reason)}`)),
      // Its own messages ("Could not parse HEIF file") are not the pane's business.
      print: () => undefined,
      printErr: () => undefined,
    });
    settle();
  });
}

/** The first picture of a HEIC/HEIF file, upright (libheif applies the file's rotation), as RGBA. */
export async function decodeHeic(
  lib: Libheif,
  bytes: Uint8Array,
): Promise<{ width: number; height: number; data: Uint8ClampedArray }> {
  const images = new lib.HeifDecoder().decode(bytes);
  try {
    const image = images[0];
    if (!image) throw new Error('This is not a HEIC photo the computer can read.');
    const width = image.get_width();
    const height = image.get_height();
    const data = await new Promise<Uint8ClampedArray>((resolve, reject) =>
      image.display({ data: new Uint8ClampedArray(width * height * 4), width, height }, (shown) =>
        shown ? resolve(shown.data) : reject(new Error('The HEIC photo could not be decoded.')),
      ),
    );
    return { width, height, data };
  } finally {
    for (const image of images) image.free();
  }
}

let loading: Promise<Libheif> | null = null;

/**
 * Loads `libheif.js` from the pane's own origin with a script element (it defines a `libheif`
 * global) and its `libheif.wasm`, and starts it; once a pane.
 */
export function loadLibheif(doc: Document): Promise<Libheif> {
  loading ??= new Promise<LibheifFactory>((resolve, reject) => {
    const win = doc.defaultView as (Window & { libheif?: LibheifFactory }) | null;
    if (win?.libheif) return resolve(win.libheif);
    const script = doc.createElement('script');
    script.src = new URL(`${LIBHEIF_FOLDER}libheif.js`, doc.baseURI).href;
    script.async = true;
    script.onload = () =>
      win?.libheif ? resolve(win.libheif) : reject(new Error('The HEIC decoder did not load.'));
    script.onerror = () => reject(new Error('The HEIC decoder could not be loaded.'));
    doc.head.appendChild(script);
  })
    .then(async (factory) => {
      const folder = new URL(LIBHEIF_FOLDER, doc.baseURI).href;
      const response = await fetch(`${folder}libheif.wasm`, { credentials: 'omit' });
      if (!response.ok) throw new Error('The HEIC decoder could not be loaded.');
      return startLibheif(factory, folder, new Uint8Array(await response.arrayBuffer()));
    })
    .catch((error: unknown) => {
      // A later photo tries again (a dropped connection, say).
      loading = null;
      throw error;
    });
  return loading;
}
