import { describe, expect, it } from 'vitest';
import { blobBytes } from '../src/office/async';
import { MAX_ATTACHMENT_BYTES, canReplyWithFile, openMail } from '../src/office/mail';
import { fakeOffice, type FakeAttachment } from './fake-office';

const bytes = (text: string) => new TextEncoder().encode(text);
const attachments: FakeAttachment[] = [
  { id: 'a1', name: 'IMG_0001.HEIC', bytes: bytes('heic') },
  { id: 'a2', name: 'logo.png', bytes: bytes('png'), isInline: true },
  { id: 'a3', name: 'Forwarded.eml', bytes: bytes('mail'), type: 'item' },
  { id: 'a4', name: 'Shared deck.pptx', bytes: bytes('link'), type: 'cloud' },
  { id: 'a5', name: 'Report.docx', bytes: bytes('docx') },
];

describe('Outlook, read mode', () => {
  it('lists only the file attachments, without inline pictures, items or cloud links', async () => {
    const { Office } = fakeOffice({ host: 'Outlook', mail: { mode: 'read', attachments } });
    const mail = openMail(Office)!;
    expect(mail.mode).toBe('read');
    expect(await mail.list()).toEqual([
      { id: 'a1', name: 'IMG_0001.HEIC', size: 4 },
      { id: 'a5', name: 'Report.docx', size: 4 },
    ]);
  });

  it("reads an attachment's bytes from its Base64 content", async () => {
    const { Office, state } = fakeOffice({ host: 'Outlook', mail: { mode: 'read', attachments } });
    const blob = await openMail(Office)!.content({ id: 'a1', name: 'IMG_0001.HEIC', size: 4 });
    expect(new TextDecoder().decode(await blobBytes(blob))).toBe('heic');
    expect(state.calls).toContain('getAttachmentContentAsync a1');
  });

  it('refuses what is not a file', async () => {
    const { Office } = fakeOffice({ host: 'Outlook', mail: { mode: 'read', attachments } });
    await expect(
      openMail(Office)!.content({ id: 'a4', name: 'Shared deck.pptx', size: 4 }),
    ).rejects.toMatchObject({ code: 'notAFile' });
  });

  it('opens a reply with the converted file (Mailbox 1.15)', async () => {
    const { Office, state } = fakeOffice({ host: 'Outlook', mail: { mode: 'read', attachments } });
    expect(canReplyWithFile(Office)).toBe(true);
    await openMail(Office)!.reply(new Blob([bytes('jpg!')]), 'IMG_0001.jpg');
    expect(state.replies).toEqual([{ name: 'IMG_0001.jpg', bytes: bytes('jpg!') }]);
  });

  it('says so where Outlook cannot attach a file to a reply', async () => {
    const { Office, state } = fakeOffice({
      host: 'Outlook',
      sets: { Mailbox: '1.14' },
      mail: { mode: 'read', attachments },
    });
    expect(canReplyWithFile(Office)).toBe(false);
    await expect(openMail(Office)!.reply(new Blob(['x']), 'x.jpg')).rejects.toMatchObject({
      code: 'unsupported',
    });
    expect(state.calls).not.toContain('displayReplyFormAsync');
  });

  it('cannot attach to a received message', async () => {
    const { Office } = fakeOffice({ host: 'Outlook', mail: { mode: 'read', attachments } });
    await expect(openMail(Office)!.attach(new Blob(['x']), 'x.jpg')).rejects.toMatchObject({
      code: 'readMode',
    });
  });
});

describe('Outlook, compose mode', () => {
  it('attaches a result and removes the original', async () => {
    const { Office, state } = fakeOffice({
      host: 'Outlook',
      mail: { mode: 'compose', attachments: attachments.slice(0, 2) },
    });
    const mail = openMail(Office)!;
    expect(mail.mode).toBe('compose');
    let changes = 0;
    mail.onChange(() => (changes += 1));
    const id = await mail.attach(new Blob([bytes('jpeg')]), 'IMG_0001.jpg');
    await mail.remove('a1');
    expect(id).toBe('att100');
    expect((await mail.list()).map((a) => a.name)).toEqual(['IMG_0001.jpg']);
    expect(state.attachments.find((a) => a.id === id)?.bytes).toEqual(bytes('jpeg'));
    await new Promise((r) => setTimeout(r, 5));
    expect(changes).toBe(2);
  });

  it("refuses a file above Outlook's 25 MB before sending it", async () => {
    const { Office, state } = fakeOffice({
      host: 'Outlook',
      mail: { mode: 'compose', attachments: [] },
    });
    const big = new Blob([new Uint8Array(MAX_ATTACHMENT_BYTES + 1)]);
    await expect(openMail(Office)!.attach(big, 'big.pdf')).rejects.toMatchObject({
      code: 'tooLarge',
    });
    expect(state.calls.some((c) => c.startsWith('addFileAttachmentFromBase64Async'))).toBe(false);
  });

  it('passes on what Outlook says when a call fails', async () => {
    const { Office } = fakeOffice({ host: 'Outlook', mail: { mode: 'compose', attachments: [] } });
    await expect(openMail(Office)!.remove('missing')).rejects.toMatchObject({
      message: 'The attachment ID was invalid.',
      code: 9002,
    });
  });

  it('has no message outside a mail item', () => {
    const { Office } = fakeOffice({ host: 'Outlook' });
    expect(openMail(Office)).toBeNull();
  });
});
