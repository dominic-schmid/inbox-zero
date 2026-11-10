/** @vitest-environment jsdom */

import React from "react";
import {
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
} from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CreateFolderDialog } from "./CreateFolderDialog";

(globalThis as { React?: typeof React }).React = React;

const { mockCreateFolderAction } = vi.hoisted(() => ({
  mockCreateFolderAction: vi.fn(),
}));
vi.mock("@/utils/actions/mail", () => ({
  createFolderAction: mockCreateFolderAction,
}));

afterEach(cleanup);

describe("CreateFolderDialog", () => {
  it("creates the folder without submitting a form it is opened from", async () => {
    mockCreateFolderAction.mockResolvedValue({
      data: { id: "new-id", displayName: "Receipts", childFolders: [] },
    });
    const onOuterSubmit = vi.fn((event: React.FormEvent) =>
      event.preventDefault(),
    );
    const onCreated = vi.fn();
    render(
      <form onSubmit={onOuterSubmit}>
        <CreateFolderDialog
          open
          onOpenChange={vi.fn()}
          emailAccountId="account-1"
          folders={[]}
          onCreated={onCreated}
        />
      </form>,
    );

    fireEvent.change(screen.getByLabelText("Name"), {
      target: { value: "Receipts" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Create" }));

    await waitFor(() =>
      expect(onCreated).toHaveBeenCalledWith({
        id: "new-id",
        path: "Receipts",
      }),
    );
    expect(onOuterSubmit).not.toHaveBeenCalled();
  });

  it("explains a name with the level separator instead of sending it", () => {
    render(
      <CreateFolderDialog
        open
        onOpenChange={vi.fn()}
        emailAccountId="account-1"
        folders={[]}
        onCreated={vi.fn()}
      />,
    );

    fireEvent.change(screen.getByLabelText("Name"), {
      target: { value: "Bills ⦙ Phone" },
    });

    expect(screen.getByText(/it separates folder levels/)).toBeTruthy();
    expect(
      (screen.getByRole("button", { name: "Create" }) as HTMLButtonElement)
        .disabled,
    ).toBe(true);
  });
});
