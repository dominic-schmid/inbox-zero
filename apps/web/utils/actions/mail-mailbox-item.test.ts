import { beforeEach, describe, expect, it, vi } from "vitest";
import { getMockEmailAccountWithAccount } from "@/__tests__/helpers";
import prisma from "@/utils/__mocks__/prisma";
import {
  createFolderAction,
  deleteMailboxItemAction,
  moveFolderAction,
  updateMailboxItemAction,
} from "@/utils/actions/mail";

vi.mock("@/utils/prisma");
vi.mock("@/utils/auth", () => ({
  auth: vi.fn(async () => ({
    user: { id: "user-1", email: "user@example.com" },
  })),
}));

const {
  mockCreateEmailProvider,
  mockCreateFolder,
  mockGetFolders,
  mockMoveFolder,
  mockDeleteFolder,
  mockDeleteLabel,
  mockRenameFolder,
  mockUpdateLabel,
} = vi.hoisted(() => ({
  mockCreateEmailProvider: vi.fn(),
  mockCreateFolder: vi.fn(),
  mockGetFolders: vi.fn(),
  mockMoveFolder: vi.fn(),
  mockDeleteFolder: vi.fn(),
  mockDeleteLabel: vi.fn(),
  mockRenameFolder: vi.fn(),
  mockUpdateLabel: vi.fn(),
}));

vi.mock("@/utils/email/provider", () => ({
  createEmailProvider: (...args: Parameters<typeof mockCreateEmailProvider>) =>
    mockCreateEmailProvider(...args),
}));

const EMAIL_ACCOUNT_ID = "email-account-1";

describe("mailbox item actions", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGetFolders.mockResolvedValue(FOLDER_TREE);
    prisma.action.findMany.mockResolvedValue([]);
    prisma.emailAccount.findUnique.mockResolvedValue(
      getMockEmailAccountWithAccount({
        email: "user@example.com",
        userId: "user-1",
        provider: "google",
      }),
    );
    mockCreateEmailProvider.mockResolvedValue({
      createFolder: mockCreateFolder,
      getFolders: mockGetFolders,
      moveFolder: mockMoveFolder,
      deleteFolder: mockDeleteFolder,
      deleteLabel: mockDeleteLabel,
      renameFolder: mockRenameFolder,
      updateLabel: mockUpdateLabel,
    });
  });

  it("renames a Gmail label through the provider", async () => {
    const result = await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "label-1",
      name: "Receipts",
      color: {
        backgroundColor: "#e66550",
        textColor: "#000000",
      },
    });

    expect(result?.serverError).toBeUndefined();
    expect(mockUpdateLabel).toHaveBeenCalledWith("label-1", {
      name: "Receipts",
      color: {
        backgroundColor: "#e66550",
        textColor: "#000000",
      },
    });
  });

  it("renames an Outlook folder through the provider", async () => {
    setProvider("microsoft");

    const result = await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "folder",
      id: "folder-1",
      name: "Projects",
    });

    expect(result?.serverError).toBeUndefined();
    expect(mockRenameFolder).toHaveBeenCalledWith("folder-1", "Projects");
  });

  it("updates an Outlook category color through the provider", async () => {
    setProvider("microsoft");

    const result = await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "category-1",
      color: {
        backgroundColor: "#1ABC9C",
        textColor: "#000000",
      },
    });

    expect(result?.serverError).toBeUndefined();
    expect(mockUpdateLabel).toHaveBeenCalledWith("category-1", {
      color: {
        backgroundColor: "#1ABC9C",
        textColor: "#000000",
      },
    });
  });

  it.each([
    ["google" as const, "#123456", "Select a supported Gmail label color."],
    [
      "microsoft" as const,
      "#654321",
      "Select a supported Outlook category color.",
    ],
  ])("rejects unsupported %s label colors before calling the provider", async (provider, backgroundColor, expectedError) => {
    setProvider(provider);

    const result = await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "label-1",
      color: { backgroundColor, textColor: "#000000" },
    });

    expect(result?.serverError).toBe(expectedError);
    expect(mockUpdateLabel).not.toHaveBeenCalled();
  });

  it.each([
    ["label" as const, mockDeleteLabel],
    ["folder" as const, mockDeleteFolder],
  ])("deletes a user-created %s through the provider", async (kind, method) => {
    if (kind === "folder") setProvider("microsoft");

    const result = await deleteMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind,
      id: `${kind}-1`,
    });

    expect(result?.serverError).toBeUndefined();
    expect(method).toHaveBeenCalledWith(`${kind}-1`);
    if (kind === "label") {
      expect(prisma.$transaction).toHaveBeenCalled();
    } else {
      expect(prisma.$executeRaw).not.toHaveBeenCalled();
    }
  });

  it("retries saved-view cleanup after a successful provider deletion", async () => {
    prisma.$transaction.mockRejectedValueOnce(
      new Error("Database unavailable"),
    );

    const firstResult = await deleteMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "label-1",
    });
    const retryResult = await deleteMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "label-1",
    });

    expect(firstResult?.serverError).toBe(
      "Failed to delete label. Please try again.",
    );
    expect(retryResult?.serverError).toBeUndefined();
    expect(mockDeleteLabel).toHaveBeenCalledTimes(2);
    expect(prisma.$transaction).toHaveBeenCalledTimes(2);
  });

  it("keeps saved views when provider deletion fails", async () => {
    mockDeleteLabel.mockRejectedValueOnce(new Error("Provider unavailable"));

    const result = await deleteMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "label",
      id: "label-1",
    });

    expect(result?.serverError).toBe(
      "Failed to delete label. Please try again.",
    );
    expect(prisma.$executeRaw).not.toHaveBeenCalled();
  });

  it("rejects folder mutations for non-Outlook accounts", async () => {
    const result = await deleteMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "folder",
      id: "folder-1",
    });

    expect(result?.serverError).toBe(
      "Folder actions are only available for Outlook accounts.",
    );
    expect(mockDeleteFolder).not.toHaveBeenCalled();
  });

  it("creates an Outlook subfolder with its name taken literally", async () => {
    setProvider("microsoft");
    mockCreateFolder.mockResolvedValue({
      id: "new-id",
      displayName: "Receipts/2024",
      childFolders: [],
    });

    const result = await createFolderAction(EMAIL_ACCOUNT_ID, {
      name: " Receipts/2024 ",
      parentId: "bills-id",
    });

    expect(result?.data?.id).toBe("new-id");
    expect(mockCreateFolder).toHaveBeenCalledWith("Receipts/2024", "bills-id");
  });

  it("moves an Outlook folder to the top level when no destination is given", async () => {
    setProvider("microsoft");

    const result = await moveFolderAction(EMAIL_ACCOUNT_ID, { id: "phone-id" });

    expect(result?.serverError).toBeUndefined();
    expect(mockMoveFolder).toHaveBeenCalledWith("phone-id", undefined);
  });

  it("refuses to move a folder into one of its own subfolders", async () => {
    setProvider("microsoft");

    const result = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "bills-id",
      destinationId: "phone-id",
    });

    expect(result?.serverError).toBe(
      "A folder can't be moved into itself or one of its subfolders.",
    );
    expect(mockMoveFolder).not.toHaveBeenCalled();
  });

  it("refuses to move a folder into Deleted Items", async () => {
    setProvider("microsoft");

    const result = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "phone-id",
      destinationId: "trash-id",
    });

    expect(result?.serverError).toBe("Folders can't be moved there.");
    expect(mockMoveFolder).not.toHaveBeenCalled();
  });

  it("reports a name clash only when Outlook says the folder exists", async () => {
    setProvider("microsoft");
    mockMoveFolder.mockRejectedValueOnce(
      Object.assign(new Error("exists"), { code: "ErrorFolderExists" }),
    );
    mockMoveFolder.mockRejectedValueOnce(new Error("timeout"));

    const clash = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "phone-id",
      destinationId: "work-id",
    });
    const other = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "phone-id",
      destinationId: "work-id",
    });

    expect(clash?.serverError).toBe(
      "A folder with that name already exists there.",
    );
    expect(other?.serverError).toBe("Failed to move folder. Please try again.");
  });

  it("updates the stored paths of rules filing into the moved folder", async () => {
    setProvider("microsoft");
    prisma.action.findMany.mockResolvedValue([
      { id: "action-1", folderId: "phone-id", folderName: "Bills ⦙ Phone" },
    ] as any);

    await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "bills-id",
      destinationId: "work-id",
    });

    expect(prisma.action.findMany).toHaveBeenCalledWith(
      expect.objectContaining({
        where: expect.objectContaining({
          folderId: { in: ["bills-id", "phone-id"] },
        }),
      }),
    );
    expect(prisma.action.updateMany).toHaveBeenCalledWith({
      where: {
        id: "action-1",
        emailAccountId: EMAIL_ACCOUNT_ID,
        folderId: "phone-id",
      },
      data: { folderName: "Work ⦙ Bills ⦙ Phone" },
    });
  });

  it("updates the stored paths of rules filing below a renamed folder", async () => {
    setProvider("microsoft");
    prisma.action.findMany.mockResolvedValue([
      { id: "action-1", folderId: "phone-id", folderName: "Bills ⦙ Phone" },
    ] as any);

    await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "folder",
      id: "bills-id",
      name: "Invoices",
    });

    expect(prisma.action.updateMany).toHaveBeenCalledWith(
      expect.objectContaining({ data: { folderName: "Invoices ⦙ Phone" } }),
    );
  });

  it("leaves rules filing into a folder in Deleted Items alone", async () => {
    setProvider("microsoft");

    await updateMailboxItemAction(EMAIL_ACCOUNT_ID, {
      kind: "folder",
      id: "trashed-id",
      name: "Older",
    });

    expect(mockRenameFolder).toHaveBeenCalledWith("trashed-id", "Older");
    expect(prisma.action.findMany).not.toHaveBeenCalled();
  });

  it("refuses to create a folder inside Deleted Items", async () => {
    setProvider("microsoft");

    const result = await createFolderAction(EMAIL_ACCOUNT_ID, {
      name: "Receipts",
      parentId: "trashed-id",
    });

    expect(result?.serverError).toBe("Folders can't be created there.");
    expect(mockCreateFolder).not.toHaveBeenCalled();
  });

  it.each([
    ["a system folder", "trash-id", "work-id", "That folder can't be moved."],
    [
      "a folder in Deleted Items",
      "trashed-id",
      undefined,
      "That folder can't be moved.",
    ],
    [
      "a folder to where it already is",
      "phone-id",
      "bills-id",
      "The folder is already there.",
    ],
  ])("refuses to move %s", async (_case, id, destinationId, message) => {
    setProvider("microsoft");

    const result = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id,
      destinationId,
    });

    expect(result?.serverError).toBe(message);
    expect(mockMoveFolder).not.toHaveBeenCalled();
  });

  it("refuses folder names that would read as a path", async () => {
    setProvider("microsoft");

    const result = await createFolderAction(EMAIL_ACCOUNT_ID, {
      name: "Bills ⦙ Phone",
    });

    expect(result?.validationErrors).toBeDefined();
    expect(mockCreateFolder).not.toHaveBeenCalled();
  });

  it("rejects moving a folder into itself", async () => {
    setProvider("microsoft");

    const result = await moveFolderAction(EMAIL_ACCOUNT_ID, {
      id: "phone-id",
      destinationId: "phone-id",
    });

    expect(result?.validationErrors).toBeDefined();
    expect(mockMoveFolder).not.toHaveBeenCalled();
  });
});

function setProvider(provider: "google" | "microsoft") {
  prisma.emailAccount.findUnique.mockResolvedValue(
    getMockEmailAccountWithAccount({
      email: "user@example.com",
      userId: "user-1",
      provider,
    }),
  );
}

const FOLDER_TREE = [
  outlookFolder("bills-id", "Bills", [outlookFolder("phone-id", "Phone")]),
  outlookFolder("work-id", "Work"),
  {
    ...outlookFolder("trash-id", "Deleted Items", [
      outlookFolder("trashed-id", "Old"),
    ]),
    systemType: "TRASH" as const,
  },
];

function outlookFolder(
  id: string,
  displayName: string,
  childFolders: ReturnType<typeof outlookFolder>[] = [],
): {
  id: string;
  displayName: string;
  childFolders: ReturnType<typeof outlookFolder>[];
  systemType?: "TRASH";
} {
  return { id, displayName, childFolders };
}
