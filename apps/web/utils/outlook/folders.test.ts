import { describe, it, expect, vi, beforeEach } from "vitest";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  getOrCreateOutlookFolderIdByName,
  getOutlookFolderTree,
  isOutlookFolderAvailable,
  getOutlookRootFolders,
  getOutlookChildFolders,
  addOutlookSystemFolderTypes,
  deleteOutlookFolder,
  renameOutlookFolder,
} from "./folders";
import type { OutlookClient } from "./client";
import { createTestLogger } from "@/__tests__/helpers";

const logger = createTestLogger();

// Mock the retry wrapper to just execute the function directly
vi.mock("@/utils/outlook/retry", () => ({
  withMicrosoftGraphRetry: <T>(fn: () => Promise<T>) => fn(),
  withMicrosoftGraphWriteRetry: <T>(fn: () => Promise<T>) => fn(),
}));

describe("Outlook folder mutations", () => {
  it("renames the selected folder", async () => {
    const patch = vi.fn().mockResolvedValue(undefined);
    const api = vi.fn().mockReturnValue({ patch });
    const client = {
      getClient: () => ({ api }),
    } as unknown as OutlookClient;

    await renameOutlookFolder(client, "folder/id", "Projects", logger);

    expect(api).toHaveBeenCalledWith("/me/mailFolders/folder%2Fid");
    expect(patch).toHaveBeenCalledWith({ displayName: "Projects" });
  });

  it("deletes the selected folder", async () => {
    const deleteFolder = vi.fn().mockResolvedValue(undefined);
    const api = vi.fn().mockReturnValue({ delete: deleteFolder });
    const client = {
      getClient: () => ({ api }),
    } as unknown as OutlookClient;

    await deleteOutlookFolder(client, "folder/id", logger);

    expect(api).toHaveBeenCalledWith("/me/mailFolders/folder%2Fid");
    expect(deleteFolder).toHaveBeenCalledOnce();
  });
});

describe("getOrCreateOutlookFolderIdByName", () => {
  it.each([
    "Bills & Receipts",
    "Partner's Invoices",
    "Ticket #42 + 50% / Q?",
    "Reçus & Factures",
  ])("finds the existing folder %s without breaking the Graph URL", async (name) => {
    const { client, requests } = createGraphClient((request) =>
      request.method === "GET"
        ? Response.json({ value: [{ id: "existing-id", displayName: name }] })
        : Response.json({ error: { code: "Unexpected" } }, { status: 500 }),
    );

    const id = await getOrCreateOutlookFolderIdByName(client, name, logger);

    expect(id).toBe("existing-id");
    expect(requests).toHaveLength(1);
    const url = new URL(requests[0].url);
    expect(url.hash).toBe("");
    expect([...url.searchParams.keys()].sort()).toEqual([
      "$filter",
      "$select",
      "$top",
    ]);
    expect(url.searchParams.get("$filter")).toBe(
      `displayName eq '${name.replaceAll("'", "''")}'`,
    );
  });

  it("uses the existing folder when creation reports it already exists", async () => {
    const name = "Bills & Receipts";
    const { client, requests } = createGraphClient((request) => {
      if (request.method === "POST") {
        return Response.json(
          { error: { code: "ErrorFolderExists", message: "exists" } },
          { status: 409 },
        );
      }
      const url = new URL(request.url);
      if (url.searchParams.has("$filter")) {
        return Response.json({ value: [] });
      }
      if (!url.searchParams.has("$skiptoken")) {
        return Response.json({
          value: [
            { id: "inbox-id", displayName: "Inbox" },
            { id: "trailing-space-id", displayName: "Bills & Receipts " },
          ],
          "@odata.nextLink":
            "https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=page2",
        });
      }
      return Response.json({
        value: [{ id: "existing-id", displayName: "bills & receipts" }],
      });
    });

    const id = await getOrCreateOutlookFolderIdByName(client, name, logger);

    expect(id).toBe("existing-id");
    expect(requests.map((request) => request.method)).toContain("POST");
  });
});

describe("getOrCreateOutlookFolderIdByName with a folder path", () => {
  it("reuses existing levels and creates the missing ones under their parent", async () => {
    const mailbox = createFakeMailbox([{ id: "bills-id", name: "Bills" }]);

    const id = await getOrCreateOutlookFolderIdByName(
      mailbox.client,
      "Bills ⦙ Phone ⦙ 2026",
      logger,
    );

    expect(mailbox.created).toEqual([
      { parent: "bills-id", name: "Phone" },
      { parent: "created-1", name: "2026" },
    ]);
    expect(id).toBe("created-2");
  });

  it("finds an existing nested folder without creating anything", async () => {
    const mailbox = createFakeMailbox([
      { id: "bills-id", name: "Bills" },
      { id: "phone-id", name: "Phone", parent: "bills-id" },
    ]);

    const id = await getOrCreateOutlookFolderIdByName(
      mailbox.client,
      "Bills ⦙ Phone",
      logger,
    );

    expect(id).toBe("phone-id");
    expect(mailbox.created).toEqual([]);
  });

  it.each([
    ["a separator character without its spaces", "A⦙B"],
    ["a separator missing one space", "Bills ⦙Phone"],
    ["a separator character inside a level", "Bills ⦙ Ph⦙one"],
    ["a leading separator", " ⦙ Bills"],
    ["a trailing separator", "Bills ⦙ "],
    ["an empty middle level", "Bills ⦙  ⦙ Phone"],
    ["only a separator", " ⦙ "],
  ])("refuses a path with %s before touching the mailbox", async (_case, path) => {
    const mailbox = createFakeMailbox([]);

    await expect(
      getOrCreateOutlookFolderIdByName(mailbox.client, path, logger),
    ).rejects.toMatchObject({ statusCode: 400 });
    expect(mailbox.requests).toEqual([]);
  });

  it("matches a nested level case-insensitively after a conflict", async () => {
    const mailbox = createFakeMailbox(
      [
        { id: "bills-id", name: "Bills" },
        { id: "phone-id", name: "PHONE", parent: "bills-id" },
      ],
      { hideFromFilter: ["phone-id"], caseInsensitiveConflicts: true },
    );

    const id = await getOrCreateOutlookFolderIdByName(
      mailbox.client,
      "Bills ⦙ Phone",
      logger,
    );

    expect(id).toBe("phone-id");
    expect(mailbox.created).toEqual([]);
  });

  it("keeps a slash as part of a top-level folder name", async () => {
    const mailbox = createFakeMailbox([]);

    await getOrCreateOutlookFolderIdByName(
      mailbox.client,
      "Receipts/2024",
      logger,
    );

    expect(mailbox.created).toEqual([{ parent: null, name: "Receipts/2024" }]);
  });

  it("uses the existing child when its creation reports it already exists", async () => {
    const mailbox = createFakeMailbox(
      [
        { id: "bills-id", name: "Bills" },
        { id: "phone-id", name: "Phone", parent: "bills-id" },
      ],
      { hideFromFilter: ["phone-id"] },
    );

    const id = await getOrCreateOutlookFolderIdByName(
      mailbox.client,
      "Bills ⦙ Phone",
      logger,
    );

    expect(id).toBe("phone-id");
  });
});

describe("isOutlookFolderAvailable", () => {
  const folders: Record<string, { id: string; parentFolderId?: string }> = {
    deleteditems: { id: "deleted-id", parentFolderId: "root-id" },
    msgfolderroot: { id: "root-id", parentFolderId: "mailbox-id" },
    "deleted-id": { id: "deleted-id", parentFolderId: "root-id" },
    "bills-id": { id: "bills-id", parentFolderId: "root-id" },
    "phone-id": { id: "phone-id", parentFolderId: "bills-id" },
    "trashed-bills-id": {
      id: "trashed-bills-id",
      parentFolderId: "deleted-id",
    },
    "trashed-phone-id": {
      id: "trashed-phone-id",
      parentFolderId: "trashed-bills-id",
    },
  };
  const serveFolders = (request: Request) => {
    const folder = folders[requestedFolderId(request)];
    return folder
      ? Response.json(folder)
      : Response.json(
          { error: { code: "ErrorItemNotFound", message: "missing" } },
          { status: 404 },
        );
  };
  const { client } = createGraphClient(serveFolders);

  it("accepts a nested folder outside Deleted Items", async () => {
    expect(await isOutlookFolderAvailable(client, "phone-id", logger)).toBe(
      true,
    );
  });

  it("rejects a folder that no longer exists", async () => {
    expect(await isOutlookFolderAvailable(client, "gone-id", logger)).toBe(
      false,
    );
  });

  it("rejects a folder deleted along with its parent", async () => {
    expect(
      await isOutlookFolderAvailable(client, "trashed-phone-id", logger),
    ).toBe(false);
  });

  it("keeps the folder when Graph can't say where it is", async () => {
    const { client: failingClient } = createGraphClient((request) =>
      requestedFolderId(request) === "msgfolderroot"
        ? Response.json(
            { error: { code: "ServiceUnavailable", message: "down" } },
            { status: 503 },
          )
        : serveFolders(request),
    );

    expect(
      await isOutlookFolderAvailable(failingClient, "phone-id", logger),
    ).toBe(true);
  });

  it("stops walking a parent chain that never reaches the mailbox root", async () => {
    const { client: loopingClient, requests } = createGraphClient((request) =>
      requestedFolderId(request) === "loop-id"
        ? Response.json({ id: "loop-id", parentFolderId: "loop-id" })
        : serveFolders(request),
    );

    expect(
      await isOutlookFolderAvailable(loopingClient, "loop-id", logger),
    ).toBe(true);
    expect(requests.length).toBeLessThan(40);
  });
});

describe("getOrCreateOutlookFolderIdByName conflict recovery", () => {
  it("rethrows the original conflict when the fallback listing fails", async () => {
    const { client } = createGraphClient((request) => {
      if (request.method === "POST") {
        return Response.json(
          { error: { code: "ErrorFolderExists", message: "exists" } },
          { status: 409 },
        );
      }
      if (new URL(request.url).searchParams.has("$filter")) {
        return Response.json({ value: [] });
      }
      return Response.json(
        { error: { code: "ServiceUnavailable", message: "down" } },
        { status: 503 },
      );
    });

    await expect(
      getOrCreateOutlookFolderIdByName(client, "Bills & Receipts", logger),
    ).rejects.toMatchObject({ code: "ErrorFolderExists" });
  });
});

function requestedFolderId(request: Request) {
  return decodeURIComponent(
    new URL(request.url).pathname.split("/mailFolders/")[1] ?? "",
  );
}

function createGraphClient(
  handler: (request: Request) => Response | Promise<Response>,
) {
  const requests: Request[] = [];
  const graph = Client.initWithMiddleware({
    middleware: {
      execute: async (context) => {
        const request = new Request(
          context.request as string,
          context.options as RequestInit,
        );
        requests.push(request);
        context.response = await handler(request);
      },
    },
  });
  const client = { getClient: () => graph } as unknown as OutlookClient;
  return { client, requests };
}

function createFakeMailbox(
  initial: { id: string; name: string; parent?: string }[],
  {
    hideFromFilter = [],
    caseInsensitiveConflicts = false,
  }: { hideFromFilter?: string[]; caseInsensitiveConflicts?: boolean } = {},
) {
  const folders = initial.map((folder) => ({
    ...folder,
    parent: folder.parent ?? null,
  }));
  const created: { parent: string | null; name: string }[] = [];
  const { client, requests } = createGraphClient(async (request) => {
    const url = new URL(request.url);
    const parentMatch = url.pathname.match(
      /mailFolders\/([^/]+)\/childFolders/,
    );
    const parent = parentMatch ? decodeURIComponent(parentMatch[1]) : null;
    const siblings = folders.filter((folder) => folder.parent === parent);

    if (request.method === "POST") {
      const { displayName } = (await request.json()) as { displayName: string };
      const sameName = (name: string) =>
        caseInsensitiveConflicts
          ? name.toLowerCase() === displayName.toLowerCase()
          : name === displayName;
      if (siblings.some((folder) => sameName(folder.name))) {
        return Response.json(
          { error: { code: "ErrorFolderExists", message: "exists" } },
          { status: 409 },
        );
      }
      const folder = {
        id: `created-${created.length + 1}`,
        name: displayName,
        parent,
      };
      folders.push(folder);
      created.push({ parent, name: displayName });
      return Response.json({ id: folder.id, displayName });
    }

    const filter = url.searchParams.get("$filter");
    const matches = filter
      ? siblings.filter(
          (folder) =>
            !hideFromFilter.includes(folder.id) &&
            filter === `displayName eq '${folder.name}'`,
        )
      : siblings;
    return Response.json({
      value: matches.map((folder) => ({
        id: folder.id,
        displayName: folder.name,
      })),
    });
  });
  return { client, created, requests };
}

function createMockClient(
  mockResponses: Record<string, { value: unknown[] }>,
): OutlookClient {
  const mockGet = vi.fn();
  const mockApi = vi.fn().mockReturnValue({
    select: vi.fn().mockReturnThis(),
    top: vi.fn().mockReturnThis(),
    expand: vi.fn().mockReturnThis(),
    get: mockGet,
  });

  mockGet.mockImplementation(() => {
    const lastCall = mockApi.mock.calls[mockApi.mock.calls.length - 1];
    const endpoint = lastCall?.[0] as string;

    // Sort patterns by length (longest first) to match more specific patterns first
    const sortedPatterns = Object.entries(mockResponses).sort(
      ([a], [b]) => b.length - a.length,
    );

    for (const [pattern, response] of sortedPatterns) {
      if (endpoint.includes(pattern)) {
        return Promise.resolve(response);
      }
    }

    return Promise.resolve({ value: [] });
  });

  return {
    getClient: () => ({
      api: mockApi,
    }),
  } as unknown as OutlookClient;
}

describe("getOutlookFolderTree", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should return root folders with their children from initial fetch", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: "inbox-id",
            displayName: "Inbox",
            childFolderCount: 1,
            childFolders: [
              {
                id: "child1-id",
                displayName: "Child1",
                childFolderCount: 0,
                childFolders: [],
              },
            ],
          },
        ],
      },
    });

    const result = await getOutlookFolderTree(mockClient, 2, logger);

    expect(result).toHaveLength(1);
    expect(result[0].displayName).toBe("Inbox");
    expect(result[0].childFolders).toHaveLength(1);
    expect(result[0].childFolders[0].displayName).toBe("Child1");
  });

  it("should recursively fetch nested folders beyond initial 2 levels", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: "inbox-id",
            displayName: "Inbox",
            childFolderCount: 1,
            childFolders: [
              {
                id: "level1-id",
                displayName: "Level1",
                childFolderCount: 1, // Has children that need fetching
                childFolders: [], // Empty - needs to be fetched
              },
            ],
          },
        ],
      },
      "level1-id/childFolders": {
        value: [
          {
            id: "level2-id",
            displayName: "Level2",
            childFolderCount: 1, // Has children
            childFolders: [], // Empty - needs to be fetched
          },
        ],
      },
      "level2-id/childFolders": {
        value: [
          {
            id: "level3-id",
            displayName: "Level3",
            childFolderCount: 0,
            childFolders: [],
          },
        ],
      },
    });

    const result = await getOutlookFolderTree(mockClient, 6, logger);

    expect(result).toHaveLength(1);
    expect(result[0].displayName).toBe("Inbox");

    const level1 = result[0].childFolders[0];
    expect(level1.displayName).toBe("Level1");

    const level2 = level1.childFolders[0];
    expect(level2.displayName).toBe("Level2");

    const level3 = level2.childFolders[0];
    expect(level3.displayName).toBe("Level3");
  });

  it("should handle multiple root folders with nested children", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: "inbox-id",
            displayName: "Inbox",
            childFolderCount: 1,
            childFolders: [
              {
                id: "inbox-child-id",
                displayName: "InboxChild",
                childFolderCount: 1, // Has children
                childFolders: [],
              },
            ],
          },
          {
            id: "drafts-id",
            displayName: "Drafts",
            childFolderCount: 0,
            childFolders: [],
          },
        ],
      },
      "inbox-child-id/childFolders": {
        value: [
          {
            id: "nested-id",
            displayName: "NestedFolder",
            childFolderCount: 0,
            childFolders: [],
          },
        ],
      },
    });

    const result = await getOutlookFolderTree(mockClient, 6, logger);

    expect(result).toHaveLength(2);

    const inbox = result.find((f) => f.displayName === "Inbox");
    expect(inbox?.childFolders[0].displayName).toBe("InboxChild");
    expect(inbox?.childFolders[0].childFolders[0].displayName).toBe(
      "NestedFolder",
    );

    const drafts = result.find((f) => f.displayName === "Drafts");
    expect(drafts?.childFolders).toHaveLength(0);
  });

  it("should respect maxDepth and not fetch beyond it", async () => {
    const apiCallTracker: string[] = [];

    const mockGet = vi.fn();
    const mockApi = vi.fn().mockImplementation((endpoint: string) => {
      apiCallTracker.push(endpoint);
      return {
        select: vi.fn().mockReturnThis(),
        top: vi.fn().mockReturnThis(),
        expand: vi.fn().mockReturnThis(),
        get: mockGet,
      };
    });

    mockGet.mockImplementation(() => {
      const lastEndpoint = apiCallTracker[apiCallTracker.length - 1];

      if (lastEndpoint === "/me/mailFolders") {
        return Promise.resolve({
          value: [
            {
              id: "root-id",
              displayName: "Root",
              childFolderCount: 1,
              childFolders: [
                {
                  id: "level1-id",
                  displayName: "Level1",
                  childFolderCount: 1, // Has children
                  childFolders: [],
                },
              ],
            },
          ],
        });
      }

      if (lastEndpoint.includes("level1-id/childFolders")) {
        return Promise.resolve({
          value: [
            {
              id: "level2-id",
              displayName: "Level2",
              childFolderCount: 1, // Has children
              childFolders: [],
            },
          ],
        });
      }

      if (lastEndpoint.includes("level2-id/childFolders")) {
        return Promise.resolve({
          value: [
            {
              id: "level3-id",
              displayName: "Level3",
              childFolderCount: 0,
              childFolders: [],
            },
          ],
        });
      }

      return Promise.resolve({ value: [] });
    });

    const mockClient = {
      getClient: () => ({
        api: mockApi,
      }),
    } as unknown as OutlookClient;

    // With maxDepth=3, should fetch root (1), level1's children (2),
    // but NOT level2's children (would be depth 3)
    await getOutlookFolderTree(mockClient, 3, logger);

    // Should NOT have called the API for level2's children
    const level2ChildCalls = apiCallTracker.filter((call) =>
      call.includes("level2-id/childFolders"),
    );
    expect(level2ChildCalls).toHaveLength(0);
  });

  it("should handle folders with no children gracefully", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: "empty-id",
            displayName: "EmptyFolder",
            childFolderCount: 0,
            childFolders: [],
          },
        ],
      },
    });

    const result = await getOutlookFolderTree(mockClient, 6, logger);

    expect(result).toHaveLength(1);
    expect(result[0].displayName).toBe("EmptyFolder");
    expect(result[0].childFolders).toHaveLength(0);
  });

  it("should only fetch children for folders with childFolderCount > 0", async () => {
    const apiCallTracker: string[] = [];

    const mockGet = vi.fn();
    const mockApi = vi.fn().mockImplementation((endpoint: string) => {
      apiCallTracker.push(endpoint);
      return {
        select: vi.fn().mockReturnThis(),
        top: vi.fn().mockReturnThis(),
        expand: vi.fn().mockReturnThis(),
        get: mockGet,
      };
    });

    mockGet.mockImplementation(() => {
      const lastEndpoint = apiCallTracker[apiCallTracker.length - 1];

      if (lastEndpoint === "/me/mailFolders") {
        return Promise.resolve({
          value: [
            {
              id: "no-children-id",
              displayName: "NoChildren",
              childFolderCount: 0, // No children
              childFolders: [],
            },
            {
              id: "has-children-id",
              displayName: "HasChildren",
              childFolderCount: 1, // Has children
              childFolders: [],
            },
          ],
        });
      }

      if (lastEndpoint.includes("has-children-id/childFolders")) {
        return Promise.resolve({
          value: [
            {
              id: "child-id",
              displayName: "Child",
              childFolderCount: 0,
              childFolders: [],
            },
          ],
        });
      }

      return Promise.resolve({ value: [] });
    });

    const mockClient = {
      getClient: () => ({
        api: mockApi,
      }),
    } as unknown as OutlookClient;

    await getOutlookFolderTree(mockClient, 6, logger);

    // Should NOT have fetched children for no-children-id
    const noChildrenCalls = apiCallTracker.filter((call) =>
      call.includes("no-children-id/childFolders"),
    );
    expect(noChildrenCalls).toHaveLength(0);

    // SHOULD have fetched children for has-children-id
    const hasChildrenCalls = apiCallTracker.filter((call) =>
      call.includes("has-children-id/childFolders"),
    );
    expect(hasChildrenCalls).toHaveLength(1);
  });
});

describe("getOutlookRootFolders", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should convert MailFolder to OutlookFolder format", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: "folder-id",
            displayName: "TestFolder",
            childFolderCount: 1,
            totalItemCount: 8,
            unreadItemCount: 3,
            childFolders: [
              {
                id: "child-id",
                displayName: "ChildFolder",
                childFolderCount: 0,
                totalItemCount: 2,
                unreadItemCount: 1,
                childFolders: [],
              },
            ],
          },
        ],
      },
    });

    const result = await getOutlookRootFolders(mockClient);

    expect(result).toEqual([
      {
        id: "folder-id",
        displayName: "TestFolder",
        childFolderCount: 1,
        totalItemCount: 8,
        unreadItemCount: 3,
        childFolders: [
          {
            id: "child-id",
            displayName: "ChildFolder",
            childFolderCount: 0,
            totalItemCount: 2,
            unreadItemCount: 1,
            childFolders: [],
          },
        ],
      },
    ]);
  });

  it("should handle null/undefined values in MailFolder", async () => {
    const mockClient = createMockClient({
      "/me/mailFolders": {
        value: [
          {
            id: null,
            displayName: undefined,
            childFolderCount: null,
            childFolders: null,
          },
        ],
      },
    });

    const result = await getOutlookRootFolders(mockClient);

    expect(result).toEqual([
      {
        id: "",
        displayName: "",
        childFolderCount: 0,
        childFolders: [],
        totalItemCount: 0,
        unreadItemCount: 0,
      },
    ]);
  });
});

describe("addOutlookSystemFolderTypes", () => {
  it("marks well-known folders by id without relying on localized names", () => {
    const [inbox, custom] = addOutlookSystemFolderTypes(
      [
        {
          id: "inbox-id",
          displayName: "Boîte de réception",
          childFolders: [],
        },
        { id: "custom-id", displayName: "Projects", childFolders: [] },
      ],
      { inbox: "inbox-id" },
    );

    expect(inbox?.systemType).toBe("INBOX");
    expect(custom?.systemType).toBeUndefined();
  });
});

describe("getOutlookChildFolders", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should fetch children for a specific folder", async () => {
    const mockClient = createMockClient({
      "parent-id/childFolders": {
        value: [
          {
            id: "child1-id",
            displayName: "Child1",
            childFolderCount: 0,
            childFolders: [],
          },
          {
            id: "child2-id",
            displayName: "Child2",
            childFolderCount: 0,
            childFolders: [],
          },
        ],
      },
    });

    const result = await getOutlookChildFolders(mockClient, "parent-id");

    expect(result).toHaveLength(2);
    expect(result[0].displayName).toBe("Child1");
    expect(result[1].displayName).toBe("Child2");
  });
});
