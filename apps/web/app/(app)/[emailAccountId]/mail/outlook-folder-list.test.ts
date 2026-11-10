import { describe, expect, it } from "vitest";
import type { OutlookFolder } from "@/utils/outlook/folders";
import { FOLDER_SEPARATOR } from "@/utils/outlook/folders";
import {
  countRulesFilingInto,
  getParentFolder,
  getMailSearchFolders,
  getMailSidebarFolders,
  moveFolderInTree,
} from "./outlook-folder-list";

function folder(
  id: string,
  childFolders: OutlookFolder[] = [],
  systemType?: OutlookFolder["systemType"],
): OutlookFolder {
  return {
    id,
    displayName: id,
    childFolders,
    childFolderCount: childFolders.length,
    totalItemCount: 0,
    unreadItemCount: 0,
    systemType,
  };
}

describe("getMailSidebarFolders", () => {
  it("keeps folders with missing child arrays without crashing", () => {
    const leaf = folder("leaf");
    Reflect.deleteProperty(leaf, "childFolders");
    const systemFolder = folder("sent", [], "SENT");
    Reflect.deleteProperty(systemFolder, "childFolders");

    const result = getMailSidebarFolders([
      folder("parent", [leaf]),
      systemFolder,
      folder("sibling"),
    ]);

    expect(result.map(({ id, depth }) => ({ id, depth }))).toEqual([
      { id: "parent", depth: 0 },
      { id: "leaf", depth: 1 },
      { id: "sibling", depth: 0 },
    ]);
  });

  it("omits duplicate system rows but keeps folders nested below them", () => {
    const result = getMailSidebarFolders([
      folder(
        "inbox-id",
        [folder("inbox-child", [folder("inbox-grandchild")])],
        "INBOX",
      ),
      folder("root-folder"),
    ]);

    expect(result.map(({ id, depth }) => ({ id, depth }))).toEqual([
      { id: "inbox-child", depth: 0 },
      { id: "inbox-grandchild", depth: 1 },
      { id: "root-folder", depth: 0 },
    ]);
  });

  it("returns no rows when Outlook only provides system folders", () => {
    expect(
      getMailSidebarFolders([
        folder("inbox-id", [], "INBOX"),
        folder("sent-id", [], "SENT"),
      ]),
    ).toEqual([]);
  });
});

describe("getMailSearchFolders", () => {
  it("lists custom folder paths and skips well-known folders", () => {
    expect(
      getMailSearchFolders([
        folder("inbox-id", [folder("Projects")], "INBOX"),
        folder("Clients"),
      ]),
    ).toEqual([
      { name: `inbox-id${FOLDER_SEPARATOR}Projects` },
      { name: "Clients" },
    ]);
  });
});

describe("getParentFolder", () => {
  const tree = [folder("bills", [folder("phone", [folder("2026")])])];

  it("returns the parent with its path, or nothing at the top level", () => {
    expect(getParentFolder(tree, "2026")).toEqual({
      id: "phone",
      path: `bills${FOLDER_SEPARATOR}phone`,
    });
    expect(getParentFolder(tree, "bills")).toBeUndefined();
  });
});

describe("moveFolderInTree", () => {
  it("moves the folder with its subfolders", () => {
    const tree = [
      folder("bills", [folder("phone", [folder("2026")])]),
      folder("work"),
    ];

    const moved = moveFolderInTree(tree, "phone", "work");

    expect(moved[0].childFolders).toEqual([]);
    expect(moved[1].childFolders.map((child) => child.id)).toEqual(["phone"]);
    expect(moved[1].childFolders[0].childFolders[0].id).toBe("2026");
  });

  it("moves a subfolder to the top level", () => {
    const tree = [folder("bills", [folder("phone")])];

    expect(moveFolderInTree(tree, "phone").map((f) => f.id)).toEqual([
      "bills",
      "phone",
    ]);
  });
});

describe("countRulesFilingInto", () => {
  it("counts rules moving mail into the folder or its subfolders once each", () => {
    const tree = [folder("bills", [folder("phone")]), folder("work")];
    const rules = [
      { actions: [{ type: "MOVE_FOLDER", folderId: "phone" }] },
      {
        actions: [
          { type: "MOVE_FOLDER", folderId: "bills" },
          { type: "MOVE_FOLDER", folderId: "phone" },
        ],
      },
      { actions: [{ type: "MOVE_FOLDER", folderId: "work" }] },
      { actions: [{ type: "LABEL", folderId: null }] },
    ];

    expect(countRulesFilingInto(tree, "bills", rules)).toBe(2);
    expect(countRulesFilingInto(tree, "phone", rules)).toBe(2);
    expect(countRulesFilingInto(tree, "missing", rules)).toBe(0);
  });
});
