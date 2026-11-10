/** @vitest-environment jsdom */

import React from "react";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { TagIcon } from "lucide-react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { OutlookFolder } from "@/utils/outlook/folders";
import { MailSidebar, type MailSidebarProps } from "./MailSidebar";

(globalThis as { React?: typeof React }).React = React;

vi.mock("next/link", () => ({
  default: ({
    children,
    href,
    onNavigate,
    ...props
  }: React.ComponentProps<"a"> & { onNavigate?: () => void }) => (
    <a href={href} onClick={() => onNavigate?.()} {...props}>
      {children}
    </a>
  ),
}));

vi.mock("@/utils/actions/mail", () => ({ createFolderAction: vi.fn() }));

const { mockRules } = vi.hoisted(() => ({
  mockRules: {
    current: [] as
      | { actions: { type: string; folderId: string | null }[] }[]
      | undefined,
  },
}));
vi.mock("@/hooks/useRules", () => ({
  useRules: () => ({ data: mockRules.current }),
}));

vi.mock("@/components/ui/sidebar", () => ({
  SidebarTrigger: () => null,
}));

afterEach(() => {
  cleanup();
  mockRules.current = [];
});

function sidebar(overrides: Partial<MailSidebarProps> = {}) {
  return (
    <MailSidebar
      activeType="inbox"
      activeLabelId={null}
      activeFolderId={null}
      hrefFor={(target) =>
        target.kind === "type" ? `/mail?type=${target.type}` : "/mail"
      }
      labels={[]}
      folders={[]}
      countsById={new Map()}
      categories={[]}
      categoryHeading="Categories"
      labelsHeading="Labels"
      labelSingular="label"
      emailAccountId="account-1"
      onNavigate={vi.fn()}
      onCompose={vi.fn()}
      onOpenSettings={vi.fn()}
      onCreateLabel={vi.fn()}
      onEditMailboxItem={vi.fn()}
      onDeleteMailboxItem={vi.fn()}
      onMoveFolder={vi.fn()}
      onFoldersChanged={vi.fn()}
      labelEditMode="name-and-color"
      labelColorOptions={[]}
      {...overrides}
    />
  );
}

function renderSidebar(overrides: Partial<MailSidebarProps> = {}) {
  return render(sidebar(overrides));
}

describe("MailSidebar", () => {
  it("reports a click on the row for the view that's already open", () => {
    const onNavigate = vi.fn();
    renderSidebar({ activeType: "inbox", onNavigate });

    fireEvent.click(screen.getByText("Inbox"));

    expect(onNavigate).toHaveBeenCalledOnce();
  });

  it("leaves the inbox out in the open and everything else behind More", () => {
    renderSidebar();

    expect(screen.getByText("Inbox").closest("a")?.getAttribute("href")).toBe(
      "/mail?type=inbox",
    );
    expect(screen.queryByText("Sent")).toBeNull();

    fireEvent.click(screen.getByRole("button", { name: "More" }));

    expect(
      [
        "Drafts",
        "Sent",
        "Archived",
        "Starred",
        "Scheduled",
        "Spam",
        "Trash",
      ].map((name) =>
        screen.getByText(name).closest("a")?.getAttribute("href"),
      ),
    ).toEqual([
      "/mail?type=draft",
      "/mail?type=sent",
      "/mail?type=archive",
      "/mail?type=starred",
      "/mail?type=scheduled",
      "/mail?type=spam",
      "/mail?type=trash",
    ]);
  });

  it("links to Discussions with its unread count only for organization members", () => {
    const { rerender } = renderSidebar();
    expect(screen.queryByText("Discussions")).toBeNull();

    rerender(sidebar({ team: { href: "/shared", unreadCount: 3 } }));
    const row = screen.getByText("Discussions").closest("a");
    expect(row?.getAttribute("href")).toBe("/shared");
    expect(row?.textContent).toContain("3");
  });

  it("opens settings from the sidebar", () => {
    const onOpenSettings = vi.fn();
    renderSidebar({ onOpenSettings });

    fireEvent.click(screen.getByRole("button", { name: "Settings" }));

    expect(onOpenSettings).toHaveBeenCalledOnce();
  });

  it("opens the More group when one of its views is showing", () => {
    renderSidebar({ activeType: "trash" });

    expect(
      screen.getByText("Trash").closest("a")?.getAttribute("aria-current"),
    ).toBe("page");
  });

  it("keeps the active view reachable in the icon rail", () => {
    renderSidebar({ activeType: "trash", collapsed: true });

    expect(screen.getByText("Trash").closest("a")?.getAttribute("href")).toBe(
      "/mail?type=trash",
    );
  });

  it("leaves the More group out of the icon rail for other views", () => {
    // Nothing in the rail can toggle the group, so a row only earns its space
    // there while its own view is open.
    renderSidebar({ activeType: "inbox", collapsed: true });

    expect(screen.queryByText("Trash")).toBeNull();
  });

  it("closes the More group in the rail once its view is left behind", () => {
    const { rerender } = renderSidebar({
      activeType: "trash",
      collapsed: true,
    });

    expect(screen.getByText("Trash")).toBeDefined();

    // The rail offers no way back out, so the group cannot outlive its view.
    rerender(sidebar({ activeType: "inbox", collapsed: true }));

    expect(screen.queryByText("Trash")).toBeNull();
  });

  it("keeps a chosen More group open across navigation when expanded", () => {
    const { rerender } = renderSidebar({ activeType: "trash" });

    rerender(sidebar({ activeType: "inbox" }));

    expect(screen.getByText("Trash")).toBeDefined();
  });

  it("closes the Categories group in the rail once its view is left behind", () => {
    const categories = [
      { name: "Promotions", type: "CATEGORY_PROMOTIONS", Icon: TagIcon },
    ];
    const { rerender } = renderSidebar({
      activeType: "CATEGORY_PROMOTIONS",
      categories,
      collapsibleCategories: true,
      collapsed: true,
    });

    expect(screen.getByText("Promotions")).toBeDefined();

    rerender(
      sidebar({
        activeType: "inbox",
        categories,
        collapsibleCategories: true,
        collapsed: true,
      }),
    );

    expect(screen.queryByText("Promotions")).toBeNull();
  });

  it("leaves the extra views out of the combined inbox", () => {
    renderSidebar({ unified: true });

    expect(screen.queryByRole("button", { name: "More" })).toBeNull();
    expect(screen.getByText("All inboxes")).toBeDefined();
  });

  it("nests a folder by dropping it onto another and refuses its own subfolder", () => {
    const onMoveFolder = vi.fn();
    renderSidebar({
      onMoveFolder,
      folders: [
        outlookFolder("bills", [outlookFolder("phone")]),
        outlookFolder("work"),
      ],
    });
    const row = (name: string) =>
      screen.getByText(name).closest("[draggable]") as HTMLElement;
    const dataTransfer = { effectAllowed: "", dropEffect: "" };

    fireEvent.dragStart(row("bills"), { dataTransfer });
    fireEvent.dragOver(row("phone"), { dataTransfer });
    fireEvent.drop(row("phone"), { dataTransfer });
    expect(onMoveFolder).not.toHaveBeenCalled();

    fireEvent.dragStart(row("phone"), { dataTransfer });
    fireEvent.dragOver(row("work"), { dataTransfer });
    fireEvent.drop(row("work"), { dataTransfer });
    expect(onMoveFolder).toHaveBeenCalledWith("phone", "work");
  });

  it("moves a subfolder to the top level when dropped on the Folders heading", () => {
    const onMoveFolder = vi.fn();
    renderSidebar({
      onMoveFolder,
      folders: [outlookFolder("bills", [outlookFolder("phone")])],
    });
    const row = (name: string) =>
      screen.getByText(name).closest("[draggable]") as HTMLElement;
    const heading = () =>
      screen.getByText("Drop here for top level").parentElement as HTMLElement;
    const dataTransfer = { effectAllowed: "", dropEffect: "" };

    fireEvent.dragStart(row("bills"), { dataTransfer });
    fireEvent.dragOver(heading(), { dataTransfer });
    fireEvent.drop(heading(), { dataTransfer });
    expect(onMoveFolder).not.toHaveBeenCalled();

    fireEvent.dragStart(row("phone"), { dataTransfer });
    fireEvent.dragOver(heading(), { dataTransfer });
    fireEvent.drop(heading(), { dataTransfer });
    expect(onMoveFolder).toHaveBeenCalledWith("phone", undefined);
  });

  it("asks before moving a folder that rules file into", () => {
    const onMoveFolder = vi.fn();
    mockRules.current = [
      { actions: [{ type: "MOVE_FOLDER", folderId: "phone" }] },
    ];
    renderSidebar({
      onMoveFolder,
      folders: [
        outlookFolder("bills", [outlookFolder("phone")]),
        outlookFolder("work"),
      ],
    });
    const row = (name: string) =>
      screen.getByText(name).closest("[draggable]") as HTMLElement;
    const dataTransfer = { effectAllowed: "", dropEffect: "" };

    fireEvent.dragStart(row("bills"), { dataTransfer });
    fireEvent.dragOver(row("work"), { dataTransfer });
    fireEvent.drop(row("work"), { dataTransfer });

    expect(onMoveFolder).not.toHaveBeenCalled();
    expect(screen.getByText(/1 rule files mail into this folder/)).toBeTruthy();
  });

  it("asks before moving a folder while the rules are still loading", () => {
    const onMoveFolder = vi.fn();
    mockRules.current = undefined;
    renderSidebar({
      onMoveFolder,
      folders: [outlookFolder("bills"), outlookFolder("work")],
    });
    const row = (name: string) =>
      screen.getByText(name).closest("[draggable]") as HTMLElement;
    const dataTransfer = { effectAllowed: "", dropEffect: "" };

    fireEvent.dragStart(row("bills"), { dataTransfer });
    fireEvent.dragOver(row("work"), { dataTransfer });
    fireEvent.drop(row("work"), { dataTransfer });

    expect(onMoveFolder).not.toHaveBeenCalled();
    expect(
      screen.getByText("Checking which rules file into this folder…"),
    ).toBeTruthy();
  });
});

function outlookFolder(id: string, childFolders: OutlookFolder[] = []) {
  return { id, displayName: id, childFolders };
}
