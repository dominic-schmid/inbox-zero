import type { MailFolder } from "@microsoft/microsoft-graph-types";
import type { OutlookClient } from "./client-types";
import type { Logger } from "@/utils/logger";
import {
  encodeODataFilter,
  escapeODataString,
} from "@/utils/outlook/odata-escape";
import { SafeError } from "@/utils/error";
import { isFolderExistsError, isNotFoundError } from "@/utils/outlook/errors";
import { resolveMicrosoftGraphNextLink } from "@/utils/outlook/page-token";
import {
  withMicrosoftGraphRetry,
  withMicrosoftGraphWriteRetry,
} from "@/utils/outlook/retry";

export type OutlookSystemFolder =
  | "INBOX"
  | "DRAFT"
  | "SENT"
  | "ARCHIVE"
  | "TRASH"
  | "SPAM"
  | "OUTBOX";

// Mail is never filed into these, so they are not offered as places to move
// or create a folder.
const NON_FILING_SYSTEM_FOLDERS = new Set<OutlookSystemFolder>([
  "DRAFT",
  "SENT",
  "TRASH",
  "SPAM",
  "OUTBOX",
]);

const OUTLOOK_FOLDER_FIELDS =
  "id,displayName,childFolderCount,totalItemCount,unreadItemCount";

// Should not use a common separator like "/|\>" as it may be used in the folder name.
// Using U+2999 as it is unlikely to appear in normal text
export const FOLDER_SEPARATOR = " ⦙ ";
export const FOLDER_SEPARATOR_CHARACTER = FOLDER_SEPARATOR.trim();

// Outlook allows deeper nesting, but a chain longer than this is not a mailbox.
const MAX_FOLDER_DEPTH = 32;

// Graph's well-known name for the parent of the top-level folders.
const MAILBOX_ROOT_FOLDER_ID = "msgfolderroot";

export type OutlookFolder = {
  id: NonNullable<MailFolder["id"]>;
  displayName: NonNullable<MailFolder["displayName"]>;
  childFolders: OutlookFolder[];
  childFolderCount?: number;
  totalItemCount?: number;
  unreadItemCount?: number;
  systemType?: OutlookSystemFolder;
};

function convertMailFolderToOutlookFolder(folder: MailFolder): OutlookFolder {
  return {
    id: folder.id ?? "",
    displayName: folder.displayName ?? "",
    childFolders:
      folder.childFolders?.map(convertMailFolderToOutlookFolder) ?? [],
    childFolderCount: folder.childFolderCount ?? 0,
    totalItemCount: folder.totalItemCount ?? 0,
    unreadItemCount: folder.unreadItemCount ?? 0,
  };
}

export async function getOutlookRootFolders(
  client: OutlookClient,
  logger: Logger,
): Promise<OutlookFolder[]> {
  const response: { value: MailFolder[] } = await withMicrosoftGraphRetry(
    () =>
      client
        .getClient()
        .api("/me/mailFolders")
        .select(OUTLOOK_FOLDER_FIELDS)
        .top(999)
        .expand(
          `childFolders($select=${OUTLOOK_FOLDER_FIELDS};$top=999;$expand=childFolders($select=${OUTLOOK_FOLDER_FIELDS};$top=999))`,
        )
        .get(),
    logger,
  );

  return response.value.map(convertMailFolderToOutlookFolder);
}

export async function getOutlookChildFolders(
  client: OutlookClient,
  folderId: string,
  logger: Logger,
): Promise<OutlookFolder[]> {
  const response: { value: MailFolder[] } = await withMicrosoftGraphRetry(
    () =>
      client
        .getClient()
        .api(`/me/mailFolders/${folderId}/childFolders`)
        .select(OUTLOOK_FOLDER_FIELDS)
        .top(999)
        .expand(
          `childFolders($select=${OUTLOOK_FOLDER_FIELDS};$top=999;$expand=childFolders($select=${OUTLOOK_FOLDER_FIELDS};$top=999))`,
        )
        .get(),
    logger,
  );

  return response.value.map(convertMailFolderToOutlookFolder);
}

async function findOutlookFolderByName(
  client: OutlookClient,
  folderName: string,
  parentId: string | undefined,
  logger: Logger,
): Promise<OutlookFolder | undefined> {
  try {
    const response: { value: MailFolder[] } = await withMicrosoftGraphRetry(
      () =>
        client
          .getClient()
          .api(getFolderCollectionPath(parentId))
          .filter(
            encodeODataFilter(
              `displayName eq '${escapeODataString(folderName)}'`,
            ),
          )
          .select("id,displayName")
          .top(1)
          .get(),
      logger,
    );

    if (response.value && response.value.length > 0) {
      return convertMailFolderToOutlookFolder(response.value[0]);
    }
    return;
  } catch (error) {
    logger.warn("Error finding folder by name", { folderName, error });
    return;
  }
}

async function expandFolderChildren(
  client: OutlookClient,
  folder: OutlookFolder,
  currentDepth: number,
  maxDepth: number,
  logger: Logger,
): Promise<void> {
  if (currentDepth >= maxDepth) {
    return;
  }

  // Use childFolderCount to know if folder has children that need fetching
  const hasUnfetchedChildren =
    (folder.childFolderCount ?? 0) > 0 &&
    (!folder.childFolders || folder.childFolders.length === 0);

  if (hasUnfetchedChildren) {
    try {
      folder.childFolders = await getOutlookChildFolders(
        client,
        folder.id,
        logger,
      );
    } catch (error) {
      logger.warn("Failed to fetch folder children", {
        folderId: folder.id,
        folderName: folder.displayName,
        error,
      });
      return;
    }
  }

  // Recursively expand children
  for (const child of folder.childFolders || []) {
    await expandFolderChildren(
      client,
      child,
      currentDepth + 1,
      maxDepth,
      logger,
    );
  }
}

export async function getOutlookFolderTree(
  client: OutlookClient,
  maxDepth: number | undefined,
  logger: Logger,
): Promise<OutlookFolder[]> {
  const folders = await getOutlookRootFolders(client, logger);

  // Recursively expand folders that have children
  // Process in parallel batches for better performance
  const expandPromises: Promise<void>[] = [];

  for (const folder of folders) {
    // Expand root folder itself if it has unfetched children
    expandPromises.push(
      expandFolderChildren(client, folder, 1, maxDepth ?? 6, logger),
    );
  }

  await Promise.all(expandPromises);

  return folders;
}

export function addOutlookSystemFolderTypes(
  folders: OutlookFolder[],
  folderIds: Record<string, string>,
): OutlookFolder[] {
  const systemTypeById = new Map<string, OutlookSystemFolder>([
    [folderIds.inbox, "INBOX"],
    [folderIds.drafts, "DRAFT"],
    [folderIds.sentitems, "SENT"],
    [folderIds.archive, "ARCHIVE"],
    [folderIds.deleteditems, "TRASH"],
    [folderIds.junkemail, "SPAM"],
    [folderIds.outbox, "OUTBOX"],
  ]);

  return folders.map((folder) => ({
    ...folder,
    systemType: systemTypeById.get(folder.id),
    childFolders: addOutlookSystemFolderTypes(folder.childFolders, folderIds),
  }));
}

/**
 * Resolves a folder name or a FOLDER_SEPARATOR path ("Bills ⦙ Phone"),
 * creating any missing level. Every other character, "/" included, is part of
 * a folder name.
 */
export async function getOrCreateOutlookFolderIdByName(
  client: OutlookClient,
  folderName: string,
  logger: Logger,
): Promise<string> {
  let parentId: string | undefined;
  for (const segment of splitOutlookFolderPath(folderName)) {
    parentId = await getOrCreateOutlookFolderId(
      client,
      segment,
      parentId,
      logger,
    );
  }
  // splitOutlookFolderPath always returns at least one segment.
  return parentId!;
}

/** Outbox is not among the cached well-known folders; it is only needed to label the tree. */
/**
 * Whether mail can still be filed into a folder known by id: it exists and is
 * not in Deleted Items, where a deleted folder lands with its subfolders.
 * When Graph can't answer, the folder counts as available: resolving the name
 * again would create a second folder next to one that is most likely fine.
 */
export async function isOutlookFolderAvailable(
  client: OutlookClient,
  folderId: string,
  logger: Logger,
): Promise<boolean> {
  const getFolder = async (id: string) => {
    try {
      const folder: MailFolder = await withMicrosoftGraphRetry(
        () =>
          client
            .getClient()
            .api(`/me/mailFolders/${encodeURIComponent(id)}`)
            .select("id,parentFolderId")
            .get(),
        logger,
      );
      return folder;
    } catch (error) {
      if (isNotFoundError(error)) return null;
      throw error;
    }
  };

  try {
    const [folder, deletedItems, root] = await Promise.all([
      getFolder(folderId),
      getFolder("deleteditems"),
      getFolder("msgfolderroot"),
    ]);
    if (!folder) return false;

    let current: MailFolder | null = folder;
    for (let depth = 0; current && depth < MAX_FOLDER_DEPTH; depth++) {
      if (current.id === deletedItems?.id) return false;
      if (!current.parentFolderId || current.parentFolderId === root?.id)
        return true;
      current = await getFolder(current.parentFolderId);
    }
    return true;
  } catch (error) {
    logger.warn("Could not check folder, keeping it", { folderId, error });
    return true;
  }
}

/** Without a parent the folder is created at the top level of the mailbox. */
export async function createOutlookFolder(
  client: OutlookClient,
  { name, parentId }: { name: string; parentId?: string },
  logger: Logger,
): Promise<OutlookFolder> {
  const folder: MailFolder = await withMicrosoftGraphWriteRetry(
    () =>
      client
        .getClient()
        .api(getFolderCollectionPath(parentId))
        .post({ displayName: name }),
    logger,
  );
  return convertMailFolderToOutlookFolder(folder);
}

/** Without a destination the folder moves to the top level of the mailbox. */
export async function moveOutlookFolder(
  client: OutlookClient,
  { folderId, destinationId }: { folderId: string; destinationId?: string },
  logger: Logger,
): Promise<void> {
  await withMicrosoftGraphWriteRetry(
    () =>
      client
        .getClient()
        .api(`/me/mailFolders/${encodeURIComponent(folderId)}/move`)
        .post({ destinationId: destinationId ?? MAILBOX_ROOT_FOLDER_ID }),
    logger,
  );
}

export async function renameOutlookFolder(
  client: OutlookClient,
  folderId: string,
  name: string,
  logger: Logger,
): Promise<void> {
  await withMicrosoftGraphWriteRetry(
    () =>
      client
        .getClient()
        .api(`/me/mailFolders/${encodeURIComponent(folderId)}`)
        .patch({ displayName: name }),
    logger,
  );
}

export async function deleteOutlookFolder(
  client: OutlookClient,
  folderId: string,
  logger: Logger,
): Promise<void> {
  await withMicrosoftGraphWriteRetry(
    () =>
      client
        .getClient()
        .api(`/me/mailFolders/${encodeURIComponent(folderId)}`)
        .delete(),
    logger,
  );
}

/** The tree without the system folders mail is never filed into, or their subfolders. */
export function withoutNonFilingOutlookFolders(
  folders: OutlookFolder[],
): OutlookFolder[] {
  return folders
    .filter(
      (folder) =>
        !folder.systemType || !NON_FILING_SYSTEM_FOLDERS.has(folder.systemType),
    )
    .map((folder) => ({
      ...folder,
      childFolders: withoutNonFilingOutlookFolders(folder.childFolders ?? []),
    }));
}

/** The folder and every folder below it. */
export function getOutlookFolderSubtree(
  folders: OutlookFolder[],
  folderId: string,
): OutlookFolder[] {
  for (const folder of folders) {
    if (folder.id === folderId) return flattenOutlookFolders([folder]);
    const subtree = getOutlookFolderSubtree(
      folder.childFolders ?? [],
      folderId,
    );
    if (subtree.length) return subtree;
  }
  return [];
}

/**
 * Whether mail and folders may be filed into the folder: it exists and is not
 * in Deleted Items, Junk, Drafts, Sent or Outbox.
 */
export function isOutlookFilingTarget(
  folders: OutlookFolder[],
  folderId: string,
): boolean {
  return (
    getOutlookFolderSubtree(withoutNonFilingOutlookFolders(folders), folderId)
      .length > 0
  );
}

/**
 * Why the folder can't be moved there, or null when the move is valid and
 * changes something. Without a destination the folder moves to the top level.
 */
export function getOutlookFolderMoveProblem(
  folders: OutlookFolder[],
  folderId: string,
  destinationId?: string,
): string | null {
  const parentById = getOutlookParentIds(folders);
  if (!parentById.has(folderId)) return "That folder no longer exists.";
  const folder = getOutlookFolderSubtree(folders, folderId)[0];
  if (folder.systemType || !isOutlookFilingTarget(folders, folderId))
    return "That folder can't be moved.";
  for (let id = destinationId; id; id = parentById.get(id) ?? undefined) {
    if (id === folderId)
      return "A folder can't be moved into itself or one of its subfolders.";
  }
  if (destinationId && !isOutlookFilingTarget(folders, destinationId))
    return "Folders can't be moved there.";
  if ((parentById.get(folderId) ?? undefined) === destinationId)
    return "The folder is already there.";
  return null;
}

/** The tree without the folder and its subfolders. */
export function withoutOutlookFolder(
  folders: OutlookFolder[],
  folderId: string,
): OutlookFolder[] {
  return folders
    .filter((folder) => folder.id !== folderId)
    .map((folder) => ({
      ...folder,
      childFolders: withoutOutlookFolder(folder.childFolders ?? [], folderId),
    }));
}

export function flattenOutlookFolders(
  folders: OutlookFolder[],
  parentPath = "",
): Array<OutlookFolder & { path: string }> {
  return folders.flatMap((folder) => {
    const path = parentPath
      ? `${parentPath}${FOLDER_SEPARATOR}${folder.displayName}`
      : folder.displayName;
    return [
      { ...folder, path },
      ...flattenOutlookFolders(folder.childFolders, path),
    ];
  });
}

// Outlook folder names are unique case-insensitively among siblings.
async function getOrCreateOutlookFolderId(
  client: OutlookClient,
  folderName: string,
  parentId: string | undefined,
  logger: Logger,
): Promise<string> {
  const existingFolder = await findOutlookFolderByName(
    client,
    folderName,
    parentId,
    logger,
  );
  if (existingFolder) return existingFolder.id;

  try {
    const folder = await createOutlookFolder(
      client,
      { name: folderName, parentId },
      logger,
    );
    return folder.id;
  } catch (error) {
    // The folder can appear between the lookup and the create, or the lookup itself can fail.
    if (isFolderExistsError(error)) {
      logger.info("Folder already exists, fetching existing folder", {
        folderName,
      });
      const folder =
        (await findOutlookFolderByName(client, folderName, parentId, logger)) ??
        (await findOutlookFolderByListing(
          client,
          folderName,
          parentId,
          logger,
        ).catch((listingError) => {
          logger.warn("Error listing folders to find existing folder", {
            folderName,
            error: listingError,
          });
          return;
        }));
      if (folder) return folder.id;
    }
    throw error;
  }
}

async function findOutlookFolderByListing(
  client: OutlookClient,
  folderName: string,
  parentId: string | undefined,
  logger: Logger,
): Promise<OutlookFolder | undefined> {
  const normalizedName = folderName.toLowerCase();
  let nextLink: string | undefined;
  do {
    const link = nextLink;
    const response: { value: MailFolder[]; "@odata.nextLink"?: string } =
      await withMicrosoftGraphRetry(
        () =>
          link
            ? client.getClient().api(link).get()
            : client
                .getClient()
                .api(getFolderCollectionPath(parentId))
                .select("id,displayName")
                .top(999)
                .get(),
        logger,
      );
    const folder = response.value.find(
      (folder) => folder.displayName?.toLowerCase() === normalizedName,
    );
    if (folder) return convertMailFolderToOutlookFolder(folder);
    nextLink =
      resolveMicrosoftGraphNextLink(response["@odata.nextLink"]) ?? undefined;
  } while (nextLink);
}

function getFolderCollectionPath(parentId: string | undefined) {
  return parentId
    ? `/me/mailFolders/${encodeURIComponent(parentId)}/childFolders`
    : "/me/mailFolders";
}

/**
 * Only FOLDER_SEPARATOR, spaces included, separates levels. A "⦙" anywhere
 * else and an empty level are refused rather than guessed at: "Bills ⦙Phone"
 * is almost always a mistyped path, not a folder name.
 */
function splitOutlookFolderPath(folderName: string): string[] {
  const levels = folderName.split(FOLDER_SEPARATOR);
  const segments =
    levels.length > 1 ? levels.map((level) => level.trim()) : levels;
  if (
    segments.some((segment) => segment.includes(FOLDER_SEPARATOR_CHARACTER))
  ) {
    throw new SafeError(
      `Folder path "${folderName}" uses "${FOLDER_SEPARATOR_CHARACTER}" outside the level separator "${FOLDER_SEPARATOR}".`,
      400,
    );
  }
  if (segments.some((segment) => !segment.trim())) {
    throw new SafeError(
      `Folder path "${folderName}" has an empty level. Separate each level with "${FOLDER_SEPARATOR}".`,
      400,
    );
  }
  return segments;
}

function getOutlookParentIds(
  folders: OutlookFolder[],
  parentId: string | null = null,
  parentById = new Map<string, string | null>(),
) {
  for (const folder of folders) {
    parentById.set(folder.id, parentId);
    getOutlookParentIds(folder.childFolders ?? [], folder.id, parentById);
  }
  return parentById;
}
