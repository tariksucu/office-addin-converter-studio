import { HostError, blobToBase64, fromBase64, officeCall } from './async';

/**
 * The open message in Outlook, through the Mailbox API (requirement set 1.8, which every
 * supported Outlook has: on the web, new and classic Outlook on Windows, Outlook on Mac).
 *
 * - Read mode: the message's file attachments (`item.attachments`) and their content
 *   (`getAttachmentContentAsync`); a result goes into a reply (`displayReplyFormAsync`).
 * - Compose mode: the draft's attachments (`getAttachmentsAsync`, `getAttachmentContentAsync`),
 *   a result is attached (`addFileAttachmentFromBase64Async`) and the original can be removed
 *   (`removeAttachmentAsync`).
 *
 * Only file attachments are offered: an attached email or calendar item, a cloud attachment (a
 * OneDrive link) and inline pictures of the body are not files to convert.
 */

export interface MailAttachment {
  readonly id: string;
  readonly name: string;
  readonly size: number;
}

export type MailMode = 'read' | 'compose';

/**
 * Any attachment of a draft: a file, a picture inside the body (`inline`), an attached email or
 * event (`item`), or a link to a file in OneDrive (`cloud`, which travels as a link only).
 */
export interface DraftAttachment extends MailAttachment {
  readonly kind: 'file' | 'inline' | 'item' | 'cloud';
}

/** The largest attachment Outlook takes from an add-in: 25 MB before Base64 (Microsoft's limit). */
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

export interface Mail {
  readonly mode: MailMode;
  list(): Promise<MailAttachment[]>;
  /** Compose: every attachment of the draft, in Outlook's order (files are `list()`'s). */
  everything(): Promise<DraftAttachment[]>;
  content(attachment: MailAttachment): Promise<Blob>;
  /** Compose: adds the file to the draft, resolving the new attachment's id. */
  attach(file: Blob, name: string): Promise<string>;
  /** Compose: removes an attachment from the draft. */
  remove(id: string): Promise<void>;
  /** Read: opens a reply with the file attached. */
  reply(file: Blob, name: string): Promise<void>;
  /** Compose: calls back when the draft's attachments change. */
  onChange(listener: () => void): void;
}

type OfficeApi = typeof Office;
type Item = Office.MessageRead & Office.MessageCompose;

interface Details {
  readonly id: string;
  readonly name: string;
  readonly size: number;
  readonly attachmentType: string;
  readonly isInline: boolean;
}

function kindOf(detail: Details, office: OfficeApi): DraftAttachment['kind'] {
  const type = office.MailboxEnums.AttachmentType;
  if (detail.attachmentType === String(type.Cloud)) return 'cloud';
  if (detail.attachmentType === String(type.Item)) return 'item';
  return detail.isInline ? 'inline' : 'file';
}

function files(details: readonly Details[], office: OfficeApi): MailAttachment[] {
  return details
    .filter(
      (a) => a.attachmentType === String(office.MailboxEnums.AttachmentType.File) && !a.isInline,
    )
    .map((a) => ({ id: a.id, name: a.name, size: a.size }));
}

/** The mailbox's item as a `Mail`, or null outside Outlook or without an open message. */
export function openMail(office: OfficeApi): Mail | null {
  const item = office.context.mailbox?.item as Item | null | undefined;
  if (!item) return null;
  // Only a draft can take attachments: the read-mode item has no such method.
  const compose = typeof item.addFileAttachmentFromBase64Async === 'function';
  const mode: MailMode = compose ? 'compose' : 'read';

  return {
    mode,
    async list() {
      if (compose)
        return files(
          await officeCall<Details[]>((done) => item.getAttachmentsAsync(done as never)),
          office,
        );
      return files(item.attachments ?? [], office);
    },
    async everything() {
      if (!compose) throw new HostError('Open the message you are writing first.', 'readMode');
      const details = await officeCall<Details[]>((done) =>
        item.getAttachmentsAsync(done as never),
      );
      return details.map((a) => ({
        id: a.id,
        name: a.name,
        size: a.size,
        kind: kindOf(a, office),
      }));
    },
    async content(attachment) {
      const content = await officeCall<Office.AttachmentContent>((done) =>
        item.getAttachmentContentAsync(attachment.id, done as never),
      );
      if (String(content.format) !== String(office.MailboxEnums.AttachmentContentFormat.Base64))
        throw new HostError(`${attachment.name} is not a file Outlook can hand over.`, 'notAFile');
      return new Blob([fromBase64(content.content)]);
    },
    async attach(file, name) {
      if (!compose) throw new HostError('Open the message you are writing first.', 'readMode');
      if (file.size > MAX_ATTACHMENT_BYTES)
        throw new HostError('Outlook takes attachments of up to 25 MB.', 'tooLarge');
      const base64 = await blobToBase64(file);
      return officeCall<string>((done) =>
        item.addFileAttachmentFromBase64Async(base64, name, done as never),
      );
    },
    async remove(id) {
      if (!compose) throw new HostError('Open the message you are writing first.', 'readMode');
      await officeCall<void>((done) => item.removeAttachmentAsync(id, done as never));
    },
    async reply(file, name) {
      if (compose) throw new HostError('This message is already being written.', 'composeMode');
      if (!office.context.requirements.isSetSupported('Mailbox', '1.15'))
        throw new HostError(
          'This version of Outlook cannot attach a file to a reply. Download it instead.',
          'unsupported',
        );
      if (file.size > MAX_ATTACHMENT_BYTES)
        throw new HostError('Outlook takes attachments of up to 25 MB.', 'tooLarge');
      const base64file = await blobToBase64(file);
      await officeCall<void>((done) =>
        item.displayReplyFormAsync(
          {
            attachments: [{ type: office.MailboxEnums.AttachmentType.Base64, name, base64file }],
          },
          done as never,
        ),
      );
    },
    onChange(listener) {
      if (!compose || typeof item.addHandlerAsync !== 'function') return;
      item.addHandlerAsync(office.EventType.AttachmentsChanged, () => listener());
    },
  };
}

/** Whether a reply can carry a converted file in this Outlook (Mailbox 1.15). */
export function canReplyWithFile(office: OfficeApi): boolean {
  return office.context.requirements.isSetSupported('Mailbox', '1.15');
}
