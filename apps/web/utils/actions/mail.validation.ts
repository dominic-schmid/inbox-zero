import { sendEmailBody } from "@/utils/types/mail";
import { labelVisibility, messageVisibility } from "@/utils/gmail/constants";
import { z } from "zod";
import { FOLDER_SEPARATOR_CHARACTER } from "@/utils/outlook/folders";

const mailboxItemKind = z.enum(["label", "folder"]);
const mailboxItemName = z
  .string()
  .trim()
  .min(1, "Mailbox item name is required")
  .max(255, "Mailbox item name must be 255 characters or fewer");

// Rule folder names nest on " ⦙ ", so a folder named with it would read as a path.
const FOLDER_NAME_PATH_CHARACTER_MESSAGE = `Folder names can't contain "${FOLDER_SEPARATOR_CHARACTER}"`;
const folderName = mailboxItemName.refine(
  (name) => !name.includes(FOLDER_SEPARATOR_CHARACTER),
  FOLDER_NAME_PATH_CHARACTER_MESSAGE,
);

export const unarchiveThreadBody = z.object({ threadId: z.string() });
export type UnarchiveThreadBody = z.infer<typeof unarchiveThreadBody>;

export const untrashThreadBody = z.object({ threadId: z.string() });
export type UntrashThreadBody = z.infer<typeof untrashThreadBody>;

export const updateMailboxItemBody = z
  .object({
    kind: mailboxItemKind,
    id: z.string().min(1, "Mailbox item ID is required"),
    name: mailboxItemName.optional(),
    color: z
      .object({
        backgroundColor: z.string().regex(/^#[0-9a-f]{6}$/i, {
          error: "Background color must be a six-digit hex color",
        }),
        textColor: z.string().regex(/^#[0-9a-f]{6}$/i, {
          error: "Text color must be a six-digit hex color",
        }),
      })
      .optional(),
    labelListVisibility: z.enum(labelVisibility).optional(),
    messageListVisibility: z.enum(messageVisibility).optional(),
  })
  .superRefine(({ kind, name, color, ...visibility }, context) => {
    const hasVisibility = Boolean(
      visibility.labelListVisibility || visibility.messageListVisibility,
    );
    if (kind === "folder" && hasVisibility) {
      context.addIssue({
        code: "custom",
        message: "Folders do not support visibility settings",
      });
    }
    if (kind === "folder" && !name) {
      context.addIssue({
        code: "custom",
        path: ["name"],
        message: "Folder name is required",
      });
    }
    if (kind === "folder" && name?.includes(FOLDER_SEPARATOR_CHARACTER)) {
      context.addIssue({
        code: "custom",
        path: ["name"],
        message: FOLDER_NAME_PATH_CHARACTER_MESSAGE,
      });
    }
    if (kind === "folder" && color) {
      context.addIssue({
        code: "custom",
        path: ["color"],
        message: "Folders do not support colors",
      });
    }
    if (kind === "label" && !name && !color && !hasVisibility) {
      context.addIssue({
        code: "custom",
        message: "A label name, color, or visibility setting is required",
      });
    }
  });

export const createFolderBody = z.object({
  name: folderName,
  /** Omitted creates a top-level folder. */
  parentId: z.string().min(1).optional(),
});

export const moveFolderBody = z
  .object({
    id: z.string().min(1, "Folder ID is required"),
    /** Omitted moves the folder to the top level. */
    destinationId: z.string().min(1).optional(),
  })
  .refine(({ id, destinationId }) => id !== destinationId, {
    message: "A folder cannot be moved into itself",
    path: ["destinationId"],
  });

export const deleteMailboxItemBody = z.object({
  kind: mailboxItemKind,
  id: z.string().min(1, "Mailbox item ID is required"),
});

export const updateDraftBody = z.object({
  draftMessageId: z.string().min(1),
  draftId: z.string().min(1).optional(),
  messageHtml: z.string().max(1_000_000),
  subject: z.string().max(10_000),
  to: z.string(),
  from: z.string().optional(),
  cc: z.string(),
  bcc: z.string(),
});

export const saveComposeDraftBody = z.object({
  draftId: z.string().min(1).optional(),
  content: sendEmailBody,
  /** Lets a send that replaces this draft find and remove it. */
  composeSessionId: z.string().min(1).max(128).optional(),
});
export const composeDraftParams = z.object({
  draftId: z.string().min(1),
});
export const readComposeDraftQuery = z.object({
  messageId: z.string().min(1),
});
