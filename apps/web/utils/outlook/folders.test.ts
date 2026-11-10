import { describe, it, expect, vi, beforeEach } from "vitest";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  getOrCreateOutlookFolderIdByName,
  getOutlookFolderTree,
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
    "Ticket #42 + 50% Q?",
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

describe("getOrCreateOutlookFolderIdByName nested paths", () => {
  it.each([
    "/ Bills // Phone /",
    "Bills/Phone/",
    "/Bills/Phone",
    "Bills//Phone",
    "Bills/Phone//",
  ])("creates Phone under Bills for %s", async (folderName) => {
    const posts: Array<{ pathname: string; body: unknown }> = [];
    const { client } = createGraphClient(async (request) => {
      const url = new URL(request.url);
      if (request.method === "POST") {
        const body = await request.json();
        posts.push({ pathname: url.pathname, body });
        const parentId = childFolderParentId(url.pathname);
        return Response.json({
          id: parentId ? "phone-id" : "bills-id",
          displayName: body.displayName,
        });
      }
      return Response.json({ value: [] });
    });

    const id = await getOrCreateOutlookFolderIdByName(
      client,
      folderName,
      logger,
    );

    expect(id).toBe("phone-id");
    expect(posts).toEqual([
      { pathname: "/v1.0/me/mailFolders", body: { displayName: "Bills" } },
      {
        pathname: "/v1.0/me/mailFolders/bills-id/childFolders",
        body: { displayName: "Phone" },
      },
    ]);
  });

  it.each([
    "bills/phone",
    "/Bills/Phone/",
    "Bills//Phone",
    " Bills / phone ",
  ])("reuses existing folders for %s", async (folderName) => {
    const { client, requests } = createGraphClient((request) => {
      const url = new URL(request.url);
      if (url.pathname.includes("/childFolders")) {
        return Response.json({
          value: [
            { id: "other-id", displayName: "Internet" },
            { id: "phone-id", displayName: "Phone" },
          ],
        });
      }
      return Response.json({
        value: [
          { id: "inbox-id", displayName: "Inbox" },
          { id: "bills-id", displayName: "Bills" },
        ],
      });
    });

    const id = await getOrCreateOutlookFolderIdByName(
      client,
      folderName,
      logger,
    );

    expect(id).toBe("phone-id");
    expect(requests.some((request) => request.method === "POST")).toBe(false);
  });

  it("handles a 409 race on a child", async () => {
    let childLookups = 0;
    const { client, requests } = createGraphClient((request) => {
      const url = new URL(request.url);
      if (request.method === "POST") {
        return Response.json(
          { error: { code: "ErrorFolderExists", message: "exists" } },
          { status: 409 },
        );
      }
      if (url.pathname.includes("/childFolders")) {
        childLookups += 1;
        if (childLookups === 1) return Response.json({ value: [] });
        return Response.json({
          value: [{ id: "phone-id", displayName: "Phone" }],
        });
      }
      return Response.json({
        value: [{ id: "bills-id", displayName: "Bills" }],
      });
    });

    const id = await getOrCreateOutlookFolderIdByName(
      client,
      "Bills/Phone",
      logger,
    );

    expect(id).toBe("phone-id");
    expect(
      requests
        .filter((request) => request.method === "POST")
        .map((request) => new URL(request.url).pathname),
    ).toEqual(["/v1.0/me/mailFolders/bills-id/childFolders"]);
  });

  it("keeps a single-segment name on the root folder lookup", async () => {
    const { client, requests } = createGraphClient((request) =>
      request.method === "GET"
        ? Response.json({
            value: [{ id: "existing-id", displayName: "Bills" }],
          })
        : Response.json({ id: "created-id", displayName: "Bills" }),
    );

    const id = await getOrCreateOutlookFolderIdByName(client, "Bills", logger);

    expect(id).toBe("existing-id");
    expect(requests).toHaveLength(1);
    const url = new URL(requests[0].url);
    expect(url.pathname).toBe("/v1.0/me/mailFolders");
    expect(url.searchParams.get("$filter")).toBe("displayName eq 'Bills'");
    expect(
      requests.some((request) => request.url.includes("childFolders")),
    ).toBe(false);
  });

  it("looks up a one-segment path as that root folder", async () => {
    const { client, requests } = createGraphClient(() =>
      Response.json({ value: [{ id: "bills-id", displayName: "Bills" }] }),
    );

    const id = await getOrCreateOutlookFolderIdByName(
      client,
      "/Bills/",
      logger,
    );

    expect(id).toBe("bills-id");
    expect(requests).toHaveLength(1);
    expect(new URL(requests[0].url).searchParams.get("$filter")).toBe(
      "displayName eq 'Bills'",
    );
  });

  it.each([
    "/",
    "///",
    " / ",
  ])("does not create a folder when %s has no names", async (folderName) => {
    const { client, requests } = createGraphClient(() =>
      Response.json({ value: [] }),
    );

    await expect(
      getOrCreateOutlookFolderIdByName(client, folderName, logger),
    ).rejects.toThrow("Outlook folder path is empty");
    expect(requests).toHaveLength(0);
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

function childFolderParentId(pathname: string): string | undefined {
  const match = pathname.match(/\/mailFolders\/([^/]+)\/childFolders$/);
  return match ? decodeURIComponent(match[1]) : undefined;
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
