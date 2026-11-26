/**
 * A JPEG's EXIF orientation: phones store a photo as the sensor saw it and say in its EXIF how to
 * turn it. Current browsers turn the picture themselves when they draw it; older webviews (and
 * any that do not) get it turned here, so a re-encoded photo is never on its side.
 */

/** The EXIF orientation (1–8) of a JPEG; 1 when it has none (or is not a JPEG). */
export function jpegOrientation(bytes: Uint8Array): number {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return 1;
  let offset = 2;
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return 1;
    const marker = bytes[offset + 1]!;
    // Fill bytes before a marker.
    if (marker === 0xff) {
      offset += 1;
      continue;
    }
    // The start of the picture data (or its end): no EXIF before it.
    if (marker === 0xda || marker === 0xd9) return 1;
    const length = (bytes[offset + 2]! << 8) | bytes[offset + 3]!;
    if (length < 2) return 1;
    if (marker === 0xe1 && isExif(bytes, offset + 4)) {
      const found = exifOrientation(
        bytes,
        offset + 10,
        Math.min(bytes.length, offset + 2 + length),
      );
      if (found) return found;
    }
    offset += 2 + length;
  }
  return 1;
}

function isExif(bytes: Uint8Array, at: number): boolean {
  const tag = [0x45, 0x78, 0x69, 0x66, 0, 0]; // "Exif\0\0"
  return tag.every((value, i) => bytes[at + i] === value);
}

/** Orientation (tag 0x0112) in the first directory of the TIFF block at `start`. */
function exifOrientation(bytes: Uint8Array, start: number, end: number): number | null {
  if (start + 8 > end) return null;
  const little = bytes[start] === 0x49 && bytes[start + 1] === 0x49; // "II"
  if (!little && !(bytes[start] === 0x4d && bytes[start + 1] === 0x4d)) return null; // "MM"
  const u16 = (at: number) =>
    little ? bytes[at]! | (bytes[at + 1]! << 8) : (bytes[at]! << 8) | bytes[at + 1]!;
  const u32 = (at: number) =>
    little
      ? (bytes[at]! | (bytes[at + 1]! << 8) | (bytes[at + 2]! << 16) | (bytes[at + 3]! << 24)) >>> 0
      : ((bytes[at]! << 24) | (bytes[at + 1]! << 16) | (bytes[at + 2]! << 8) | bytes[at + 3]!) >>>
        0;
  if (u16(start + 2) !== 42) return null;
  const directory = start + u32(start + 4);
  if (directory + 2 > end) return null;
  const entries = u16(directory);
  for (let i = 0; i < entries; i++) {
    const entry = directory + 2 + i * 12;
    if (entry + 12 > end) return null;
    if (u16(entry) === 0x0112) {
      const value = u16(entry + 8);
      return value >= 1 && value <= 8 ? value : null;
    }
  }
  return null;
}

/** Orientations 5–8 turn the picture a quarter: its width and height swap. */
export function swapsSides(orientation: number): boolean {
  return orientation >= 5 && orientation <= 8;
}

/**
 * How to draw a picture stored with `orientation` upright into a `width` × `height` canvas
 * (already the upright size): the canvas transform (`setTransform(a, b, c, d, e, f)`) and the size
 * to draw the stored picture at, before it is turned.
 */
export function uprightTransform(
  orientation: number,
  width: number,
  height: number,
): {
  matrix: [number, number, number, number, number, number];
  drawWidth: number;
  drawHeight: number;
} {
  const quarter = swapsSides(orientation);
  const drawWidth = quarter ? height : width;
  const drawHeight = quarter ? width : height;
  const matrices: Record<number, [number, number, number, number, number, number]> = {
    2: [-1, 0, 0, 1, width, 0],
    3: [-1, 0, 0, -1, width, height],
    4: [1, 0, 0, -1, 0, height],
    5: [0, 1, 1, 0, 0, 0],
    6: [0, 1, -1, 0, width, 0],
    7: [0, -1, -1, 0, width, height],
    8: [0, -1, 1, 0, 0, height],
  };
  return { matrix: matrices[orientation] ?? [1, 0, 0, 1, 0, 0], drawWidth, drawHeight };
}

/**
 * A 2 × 1 JPEG stored with orientation 6: a browser that turns photos itself shows it 1 × 2.
 * Made with sharp and a hand-written EXIF block (`MM`, one entry: 0x0112 = 6).
 */
export const ORIENTATION_PROBE =
  'data:image/jpeg;base64,/9j/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAYAAAAAAAD/2wBDABALDA4MChAODQ4SERATGCgaGBYWGDEjJR0oOjM9PDkzODdASFxOQERXRTc4UG1RV19iZ2hnPk1xeXBkeFxlZ2P/2wBDARESEhgVGC8aGi9jQjhCY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2P/wAARCAABAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAABQb/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCKALXj/9k=';
