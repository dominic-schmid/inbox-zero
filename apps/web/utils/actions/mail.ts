"use server";

import { z } from "zod";
import { markTrackedDraftDeleted } from "@/utils/ai/draft-cleanup";
import prisma from "@/utils/prisma";
import { removeLabelFromMailSplits } from "@/utils/split-inbox/splits.server";
import { sendEmailBody } from "@/utils/types/mail";
import { actionClient } from "@/utils/actions/safe-action";
import { DraftNotFoundError, SafeError } from "@/utils/error";
import { createEmailProvider } from "@/utils/email/provider";
import { sendHtmlEmailWithOpenTracking } from "@/utils/email/sent-message-open/sent-message-open.server";
import {
  createFolderBody,
  deleteMailboxItemBody,
  moveFolderBody,
  unarchiveThreadBody,
  untrashThreadBody,
  updateMailboxItemBody,
  updateDraftBody,
  saveComposeDraftBody,
  composeDraftParams,
} from "@/utils/actions/mail.validation";
import {
  isGoogleProvider,
  isMicrosoftProvider,
} from "@/utils/email/provider-types";
import { isGmailLabelColor } from "@/utils/gmail/label-colors";
import { getOutlookCategoryPreset } from "@/utils/outlook/category-colors";
import { resetDefaultLabelColors } from "@/utils/label/reset-label-colors";
import { ActionType } from "@/generated/prisma/enums";
import type { Logger } from "@/utils/logger";
import { isFolderExistsError } from "@/utils/outlook/errors";
import {
  FOLDER_SEPARATOR,
  flattenOutlookFolders,
  getOutlookFolderMoveProblem,
  getOutlookFolderSubtree,
  isOutlookFilingTarget,
  type OutlookFolder,
} from "@/utils/outlook/folders";

import {
  saveComposeDraft,
  discardComposeDraft,
} from "@/utils/email/compose-draft";

const isStatusOk = (status: number) => status >= 200 && status < 300;

export const archiveThreadAction = actionClient
  .metadata({ name: "archiveThread" })
  .inputSchema(
    z.object({ threadId: z.string(), labelId: z.string().optional() }),
  )
  .action(
    async ({
      ctx: { emailAccountId, emailAccount, provider, logger },
      parsedInput: { threadId, labelId },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        await emailProvider.archiveThreadWithLabel(
          threadId,
          emailAccount.email,
          labelId,
        );
      } catch (error) {
        logger.error("Failed to archive thread", { error });
        throw new SafeError("Failed to archive email. Please try again.");
      }
    },
  );

export const unarchiveThreadAction = actionClient
  .metadata({ name: "unarchiveThread" })
  .inputSchema(unarchiveThreadBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { threadId },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        await emailProvider.unarchiveThread(threadId);
      } catch (error) {
        logger.error("Failed to unarchive thread", { error });
        throw new SafeError("Failed to unarchive email. Please try again.");
      }
    },
  );

export const trashThreadAction = actionClient
  .metadata({ name: "trashThread" })
  .inputSchema(z.object({ threadId: z.string() }))
  .action(
    async ({
      ctx: { emailAccountId, emailAccount, provider, logger },
      parsedInput: { threadId },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        await emailProvider.trashThread(threadId, emailAccount.email, "user");
      } catch (error) {
        logger.error("Failed to trash thread", { error });
        throw new SafeError("Failed to delete email. Please try again.");
      }
    },
  );

export const untrashThreadAction = actionClient
  .metadata({ name: "untrashThread" })
  .inputSchema(untrashThreadBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { threadId },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        await emailProvider.untrashThread(threadId);
      } catch (error) {
        logger.error("Failed to untrash thread", { error });
        throw new SafeError("Failed to restore email. Please try again.");
      }
    },
  );

export const markReadThreadAction = actionClient
  .metadata({ name: "markReadThread" })
  .inputSchema(z.object({ threadId: z.string(), read: z.boolean() }))
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { threadId, read },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        await emailProvider.markReadThread(threadId, read);
      } catch (error) {
        logger.error("Failed to mark thread read state", { error });
        throw new SafeError(
          `Failed to mark email as ${read ? "read" : "unread"}. Please try again.`,
        );
      }
    },
  );

export const createAutoArchiveFilterAction = actionClient
  .metadata({ name: "createAutoArchiveFilter" })
  .inputSchema(
    z.object({
      from: z.string(),
      gmailLabelId: z.string().optional(),
      labelName: z.string().optional(),
    }),
  )
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { from, gmailLabelId, labelName },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      await emailProvider.createAutoArchiveFilter({
        from,
        gmailLabelId,
        labelName,
      });
    },
  );

export const createFilterAction = actionClient
  .metadata({ name: "createFilter" })
  .inputSchema(z.object({ from: z.string(), gmailLabelId: z.string() }))
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { from, gmailLabelId },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      const res = await emailProvider.createFilter({
        from,
        addLabelIds: [gmailLabelId],
      });

      if (!isStatusOk(res.status)) {
        logger.error("Failed to create filter", {
          from,
          gmailLabelId,
          status: res.status,
        });
        throw new SafeError("Failed to create filter");
      }
    },
  );

export const deleteFilterAction = actionClient
  .metadata({ name: "deleteFilter" })
  .inputSchema(z.object({ id: z.string() }))
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { id },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      const res = await emailProvider.deleteFilter(id);

      if (!isStatusOk(res.status)) {
        logger.error("Failed to delete filter", {
          filterId: id,
          status: res.status,
        });
        throw new SafeError("Failed to delete filter");
      }
    },
  );

export const createLabelAction = actionClient
  .metadata({ name: "createLabel" })
  .inputSchema(
    z.object({ name: z.string(), description: z.string().optional() }),
  )
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { name, description },
    }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });
      const label = await emailProvider.createLabel(name, description);
      return label;
    },
  );

export const updateMailboxItemAction = actionClient
  .metadata({ name: "updateMailboxItem" })
  .inputSchema(updateMailboxItemBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: {
        kind,
        id,
        name,
        color,
        labelListVisibility,
        messageListVisibility,
      },
    }) => {
      assertMailboxItemMutationSupported({ kind, provider });
      if (kind === "label" && isMicrosoftProvider(provider) && name) {
        throw new SafeError(
          "Outlook category names cannot be changed. Edit its color instead.",
        );
      }
      if (
        (labelListVisibility || messageListVisibility) &&
        !isGoogleProvider(provider)
      ) {
        throw new SafeError(
          "Visibility settings are only available for Gmail.",
        );
      }
      if (color && isGoogleProvider(provider) && !isGmailLabelColor(color)) {
        throw new SafeError("Select a supported Gmail label color.");
      }
      if (
        color &&
        isMicrosoftProvider(provider) &&
        !getOutlookCategoryPreset(color.backgroundColor)
      ) {
        throw new SafeError("Select a supported Outlook category color.");
      }
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        if (kind === "folder") {
          const folders = await emailProvider.getFolders();
          await emailProvider.renameFolder(id, name!);
          const renamed = flattenOutlookFolders(folders).find(
            (folder) => folder.id === id,
          );
          if (renamed) {
            await syncStoredFolderPaths({
              emailAccountId,
              folders,
              folderId: id,
              newPath:
                renamed.path.slice(0, -renamed.displayName.length) + name!,
              logger,
            });
          }
        } else
          await emailProvider.updateLabel(id, {
            name,
            color,
            labelListVisibility,
            messageListVisibility,
          });
      } catch (error) {
        logger.error("Failed to update mailbox item", { error, kind });
        throw new SafeError(`Failed to update ${kind}. Please try again.`);
      }
    },
  );

export const createFolderAction = actionClient
  .metadata({ name: "createFolder" })
  .inputSchema(createFolderBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { name, parentId },
    }) => {
      assertMailboxItemMutationSupported({ kind: "folder", provider });
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      if (
        parentId &&
        !isOutlookFilingTarget(await emailProvider.getFolders(), parentId)
      ) {
        throw new SafeError("Folders can't be created there.");
      }

      try {
        return await emailProvider.createFolder(name, parentId);
      } catch (error) {
        logger.error("Failed to create folder", { error });
        throw new SafeError(
          isFolderExistsError(error)
            ? "A folder with that name already exists there."
            : "Failed to create folder. Please try again.",
        );
      }
    },
  );

export const moveFolderAction = actionClient
  .metadata({ name: "moveFolder" })
  .inputSchema(moveFolderBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { id, destinationId },
    }) => {
      assertMailboxItemMutationSupported({ kind: "folder", provider });
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      const folders = await emailProvider.getFolders();
      const moveProblem = getOutlookFolderMoveProblem(
        folders,
        id,
        destinationId,
      );
      if (moveProblem) throw new SafeError(moveProblem);

      try {
        await emailProvider.moveFolder(id, destinationId);
      } catch (error) {
        logger.error("Failed to move folder", { error });
        throw new SafeError(
          isFolderExistsError(error)
            ? "A folder with that name already exists there."
            : "Failed to move folder. Please try again.",
        );
      }

      const flatFolders = flattenOutlookFolders(folders);
      const name = flatFolders.find((folder) => folder.id === id)!.displayName;
      const destinationPath = destinationId
        ? flatFolders.find((folder) => folder.id === destinationId)!.path
        : "";
      await syncStoredFolderPaths({
        emailAccountId,
        folders,
        folderId: id,
        newPath: joinFolderPath(destinationPath, name),
        logger,
      });
    },
  );

export const resetLabelColorsAction = actionClient
  .metadata({ name: "resetLabelColors" })
  .action(async ({ ctx: { emailAccountId, provider, logger } }) => {
    const emailProvider = await createEmailProvider({
      emailAccountId,
      provider,
      logger,
    });
    const result = await resetDefaultLabelColors({
      emailAccountId,
      emailProvider,
      logger,
    });
    if (result.failed > 0 && result.updated === 0)
      throw new SafeError("Failed to reset label colors. Please try again.");
    return result;
  });

export const deleteMailboxItemAction = actionClient
  .metadata({ name: "deleteMailboxItem" })
  .inputSchema(deleteMailboxItemBody)
  .action(
    async ({
      ctx: { emailAccountId, provider, logger },
      parsedInput: { kind, id },
    }) => {
      assertMailboxItemMutationSupported({ kind, provider });
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      try {
        if (kind === "folder") {
          await emailProvider.deleteFolder(id);
        } else {
          await emailProvider.deleteLabel(id);
          await removeLabelFromMailSplits({ emailAccountId, labelId: id });
        }
      } catch (error) {
        logger.error("Failed to delete mailbox item", { error, kind });
        throw new SafeError(`Failed to delete ${kind}. Please try again.`);
      }
    },
  );

export const updateLabelsAction = actionClient
  .metadata({ name: "updateLabels" })
  .inputSchema(
    z.object({
      labels: z.array(
        z.object({
          name: z.string(),
          description: z.string().optional(),
          enabled: z.boolean(),
          gmailLabelId: z.string(),
        }),
      ),
    }),
  )
  .action(async ({ ctx: { emailAccountId }, parsedInput: { labels } }) => {
    const enabledLabels = labels.filter((label) => label.enabled);
    const disabledLabels = labels.filter((label) => !label.enabled);

    await prisma.$transaction([
      ...enabledLabels.map((label) => {
        const { name, description, enabled, gmailLabelId } = label;

        return prisma.label.upsert({
          where: { name_emailAccountId: { name, emailAccountId } },
          create: {
            gmailLabelId,
            name,
            description,
            enabled,
            emailAccountId,
          },
          update: {
            name,
            description,
            enabled,
          },
        });
      }),
      prisma.label.deleteMany({
        where: {
          emailAccountId,
          name: { in: disabledLabels.map((label) => label.name) },
        },
      }),
    ]);
  });

export const sendEmailAction = actionClient
  .metadata({ name: "sendEmail" })
  .inputSchema(sendEmailBody)
  .action(
    async ({ ctx: { emailAccountId, provider, logger }, parsedInput }) => {
      const emailProvider = await createEmailProvider({
        emailAccountId,
        provider,
        logger,
      });

      const result = await sendHtmlEmailWithOpenTracking({
        emailAccountId,
        email: parsedInput,
        emailProvider,
        logger,
      });

      return {
        success: true,
        messageId: result.messageId,
        threadId: result.threadId,
      };
    },
  );

export const updateDraftAction = actionClient
  .metadata({ name: "updateDraft" })
  .inputSchema(updateDraftBody)
  .action(
    async ({
      ctx: { emailAccountId, provider: providerName, logger },
      parsedInput,
    }) => {
      const provider = await createEmailProvider({
        emailAccountId,
        provider: providerName,
        logger,
      });
      const draftId =
        parsedInput.draftId ??
        (await provider.getDraftReferenceForMessage(parsedInput.draftMessageId))
          ?.id;
      // A draft sent, deleted, or replaced elsewhere can never be saved, so the
      // composer is told to stop syncing rather than retry.
      if (!draftId) return { status: "missing" as const };
      const {
        draftMessageId: _messageId,
        draftId: _draftId,
        ...content
      } = parsedInput;
      try {
        await provider.updateDraft(draftId, content);
      } catch (error) {
        if (error instanceof DraftNotFoundError)
          return { status: "missing" as const };
        throw error;
      }
      // Gmail may have moved the draft to a new message; the composer needs it,
      // so a failed read fails the save and autosave retries it.
      const message = await provider.getDraft(draftId);
      return {
        status: "saved" as const,
        draftId,
        messageId: message?.id ?? null,
      };
    },
  );

function assertMailboxItemMutationSupported({
  kind,
  provider,
}: {
  kind: "label" | "folder";
  provider: string;
}) {
  if (kind === "folder" && !isMicrosoftProvider(provider)) {
    throw new SafeError(
      "Folder actions are only available for Outlook accounts.",
    );
  }
}

export const saveComposeDraftAction = actionClient
  .metadata({ name: "saveComposeDraft" })
  .inputSchema(saveComposeDraftBody)
  .action(
    async ({
      ctx: { emailAccountId, provider: providerName, logger },
      parsedInput,
    }) => {
      const provider = await createEmailProvider({
        emailAccountId,
        provider: providerName,
        logger,
      });
      const { composeSessionId, ...draft } = parsedInput;
      const saved = await saveComposeDraft({
        provider,
        ...draft,
        ...(composeSessionId
          ? { composeSession: { id: composeSessionId, emailAccountId, logger } }
          : {}),
      });
      return saved;
    },
  );

export const discardComposeDraftAction = actionClient
  .metadata({ name: "discardComposeDraft" })
  .inputSchema(composeDraftParams)
  .action(
    async ({
      ctx: { emailAccountId, provider: providerName, logger },
      parsedInput,
    }) => {
      const provider = await createEmailProvider({
        emailAccountId,
        provider: providerName,
        logger,
      });
      await discardComposeDraft({ provider, ...parsedInput });
      // An AI draft discarded here mustn't stay listed for cleanup.
      await markTrackedDraftDeleted({
        draftId: parsedInput.draftId,
        emailAccountId,
        logger,
      }).catch((error) => {
        logger.error("Failed to update tracking after deleting draft", {
          error,
        });
      });
      return { success: true };
    },
  );

/**
 * Rules file by folder id, so they follow a moved or renamed folder. The paths
 * stored on rules filing into that folder or below it are updated to match,
 * so the rule editor and the public API show where mail really goes. Other
 * rules keep their stored path, which is what a config sync compares against.
 * `folders` is the tree from before the change.
 */
async function syncStoredFolderPaths({
  emailAccountId,
  folders,
  folderId,
  newPath,
  logger,
}: {
  emailAccountId: string;
  folders: OutlookFolder[];
  folderId: string;
  newPath: string;
  logger: Logger;
}) {
  // Rules filing into Deleted Items or Junk are left alone: their next save by
  // name recreates the folder instead of filing into the trash.
  if (!isOutlookFilingTarget(folders, folderId)) return;
  const [folder, ...descendants] = getOutlookFolderSubtree(folders, folderId);
  const pathById = new Map(
    [folder, ...descendants].map((subfolder) => [
      subfolder.id,
      newPath + subfolder.path.slice(folder.displayName.length),
    ]),
  );

  try {
    const actions = await prisma.action.findMany({
      where: {
        emailAccountId,
        type: ActionType.MOVE_FOLDER,
        folderId: { in: [...pathById.keys()] },
      },
      select: { id: true, folderId: true, folderName: true },
    });
    const stale = actions.flatMap(({ id, folderId, folderName }) => {
      const path = folderId && pathById.get(folderId);
      return path && path !== folderName ? [{ id, folderId, path }] : [];
    });
    if (!stale.length) return;

    await prisma.$transaction(
      stale.map(({ id, folderId, path }) =>
        // Matching the folder read above skips a rule re-pointed meanwhile.
        prisma.action.updateMany({
          where: { id, emailAccountId, folderId },
          data: { folderName: path },
        }),
      ),
    );
  } catch (error) {
    // The folder change itself succeeded; a stale path only affects what the
    // rule shows, and the next save resolves it again.
    logger.warn("Failed to update rule folder paths", { error });
  }
}

function joinFolderPath(parentPath: string, name: string) {
  return parentPath ? `${parentPath}${FOLDER_SEPARATOR}${name}` : name;
}
