/** @vitest-environment jsdom */

import React from "react";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { MoveFolderDialog } from "./MoveFolderDialog";

(globalThis as { React?: typeof React }).React = React;

vi.mock("@/utils/actions/mail", () => ({ createFolderAction: vi.fn() }));

afterEach(cleanup);

describe("MoveFolderDialog", () => {
  it("starts where the folder is and won't move it there again", () => {
    const onMove = vi.fn();
    render(
      <MoveFolderDialog
        folder={{
          id: "phone",
          name: "Phone",
          destination: { id: "bills", path: "Bills" },
        }}
        folders={[
          {
            id: "bills",
            displayName: "Bills",
            childFolders: [
              { id: "phone", displayName: "Phone", childFolders: [] },
            ],
          },
        ]}
        ruleCount={0}
        onOpenChange={vi.fn()}
        onMove={onMove}
      />,
    );

    expect(screen.getByText("The folder is already there.")).toBeTruthy();
    const move = screen.getByRole("button", { name: "Move" });
    expect((move as HTMLButtonElement).disabled).toBe(true);
    fireEvent.click(move);
    expect(onMove).not.toHaveBeenCalled();
  });
});
