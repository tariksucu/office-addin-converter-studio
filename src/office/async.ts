/**
 * Office's callback APIs as promises. A failed call becomes a `HostError` with Office's own
 * message, which the pane shows as it is.
 */

export class HostError extends Error {
  constructor(
    message: string,
    readonly code: string | number = 'host',
  ) {
    super(message);
    this.name = 'HostError';
  }
}

interface AsyncResultLike<T> {
  readonly status: string;
  readonly value: T;
  readonly error?: { readonly message?: string; readonly code?: number | string } | null;
}

/** Calls `start` with a callback and settles once Office calls it back. */
export function officeCall<T>(
  start: (callback: (result: AsyncResultLike<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    try {
      start((result) => {
        if (result.status === 'succeeded') resolve(result.value);
        else
          reject(
            new HostError(result.error?.message || 'Office could not do it.', result.error?.code),
          );
      });
    } catch (error) {
      reject(new HostError(error instanceof Error ? error.message : String(error)));
    }
  });
}

/** Bytes as Base64, in pieces small enough for `String.fromCharCode`. */
export function toBase64(bytes: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000)
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

export function fromBase64(base64: string): Uint8Array<ArrayBuffer> {
  const binary = atob(base64.replace(/\s+/g, ''));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** A file's bytes (through `Response`, which older webviews have where `Blob.arrayBuffer` is missing). */
export async function blobBytes(blob: Blob): Promise<Uint8Array<ArrayBuffer>> {
  return new Uint8Array(await new Response(blob).arrayBuffer());
}

export async function blobToBase64(blob: Blob): Promise<string> {
  return toBase64(await blobBytes(blob));
}

export function blobText(blob: Blob): Promise<string> {
  return new Response(blob).text();
}
