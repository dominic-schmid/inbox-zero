import { ActionType } from "@/generated/prisma/enums";
import {
  flattenOutlookFolders,
  type OutlookFolder,
  withoutOutlookFolder,
} from "@/utils/outlook/folders";

export type MailSidebarFolder = OutlookFolder & { depth: number };

export function getMailSidebarFolders(
  folders: OutlookFolder[],
  depth = 0,
): MailSidebarFolder[] {
  return folders.flatMap((folder) => {
    const childDepth = folder.systemType ? depth : depth + 1;
    const children = getMailSidebarFolders(
      folder.childFolders ?? [],
      childDepth,
    );
    return folder.systemType ? children : [{ ...folder, depth }, ...children];
  });
}

export function getMailSearchFolders(
  folders: OutlookFolder[],
): { name: string }[] {
  return flattenOutlookFolders(folders)
    .filter((folder) => !folder.systemType)
    .map((folder) => ({ name: folder.path }));
}

export function moveFolderInTree(
  folders: OutlookFolder[],
  folderId: string,
  destinationId?: string,
): OutlookFolder[] {
  const moved = flattenOutlookFolders(folders).find(
    (folder) => folder.id === folderId,
  );
  if (!moved) return folders;
  const { path: _path, ...folder } = moved;
  const remaining = withoutOutlookFolder(folders, folderId);
  if (!destinationId) return [...remaining, folder];
  return insertFolder(remaining, destinationId, folder);
}

/** Rules with a move action into the folder or one of its subfolders. */
export function countRulesFilingInto(
  folders: OutlookFolder[],
  folderId: string,
  rules: { actions: { type: string; folderId: string | null }[] }[],
): number {
  const folder = flattenOutlookFolders(folders).find(
    (candidate) => candidate.id === folderId,
  );
  if (!folder) return 0;
  const folderIds = new Set(
    flattenOutlookFolders([folder]).map((candidate) => candidate.id),
  );
  return rules.filter((rule) =>
    rule.actions.some(
      (action) =>
        action.type === ActionType.MOVE_FOLDER &&
        !!action.folderId &&
        folderIds.has(action.folderId),
    ),
  ).length;
}

/** The folder's parent with its path, or undefined at the top level. */
export function getParentFolder(
  folders: OutlookFolder[],
  folderId: string,
): { id: string; path: string } | undefined {
  const parent = flattenOutlookFolders(folders).find((folder) =>
    folder.childFolders?.some((child) => child.id === folderId),
  );
  return parent && { id: parent.id, path: parent.path };
}

export function getFolderPath(
  folders: OutlookFolder[],
  folderId: string,
): string | undefined {
  return flattenOutlookFolders(folders).find((folder) => folder.id === folderId)
    ?.path;
}

function insertFolder(
  folders: OutlookFolder[],
  parentId: string,
  child: OutlookFolder,
): OutlookFolder[] {
  return folders.map((folder) =>
    folder.id === parentId
      ? {
          ...folder,
          childFolders: [...(folder.childFolders ?? []), child],
          childFolderCount: (folder.childFolderCount ?? 0) + 1,
        }
      : {
          ...folder,
          childFolders: insertFolder(
            folder.childFolders ?? [],
            parentId,
            child,
          ),
        },
  );
}
