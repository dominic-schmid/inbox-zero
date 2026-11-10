"use client";

import { type FormEvent, useState } from "react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { FolderSelector } from "@/components/FolderSelector";
import {
  getOutlookFolderMoveProblem,
  type OutlookFolder,
} from "@/utils/outlook/folders";
import {
  type FolderRuleCount,
  FolderRuleCountNotice,
} from "./FolderRuleCountNotice";

const TOP_LEVEL = { name: "", id: "" };

export type FolderMove = {
  id: string;
  name: string;
  /** Preselected: where the folder was dropped, else where it is now. */
  destination?: { id: string; path: string };
};

export function MoveFolderDialog({
  folder,
  folders,
  ruleCount,
  onOpenChange,
  onMove,
}: {
  /** The folder being moved; null closes the dialog. */
  folder: FolderMove | null;
  folders: OutlookFolder[];
  ruleCount: FolderRuleCount;
  onOpenChange: (open: boolean) => void;
  onMove: (folderId: string, destinationId?: string) => Promise<boolean>;
}) {
  return (
    <Dialog open={!!folder} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        {folder && (
          <MoveFolderForm
            folder={folder}
            folders={folders}
            ruleCount={ruleCount}
            onMove={onMove}
            onDone={() => onOpenChange(false)}
          />
        )}
      </DialogContent>
    </Dialog>
  );
}

function MoveFolderForm({
  folder,
  folders,
  ruleCount,
  onMove,
  onDone,
}: {
  folder: FolderMove;
  folders: OutlookFolder[];
  ruleCount: FolderRuleCount;
  onMove: (folderId: string, destinationId?: string) => Promise<boolean>;
  onDone: () => void;
}) {
  const [destination, setDestination] = useState(
    folder.destination
      ? { name: folder.destination.path, id: folder.destination.id }
      : TOP_LEVEL,
  );
  const [isMoving, setIsMoving] = useState(false);
  const moveProblem = getOutlookFolderMoveProblem(
    folders,
    folder.id,
    destination.id || undefined,
  );

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    event.stopPropagation();
    if (moveProblem) return;
    setIsMoving(true);
    try {
      const moved = await onMove(folder.id, destination.id || undefined);
      if (moved) onDone();
    } finally {
      setIsMoving(false);
    }
  };

  return (
    <form onSubmit={submit} className="space-y-4">
      <DialogHeader>
        <DialogTitle>Move “{folder.name}”</DialogTitle>
        <DialogDescription>
          Its subfolders move with it. Leave the destination empty to make it a
          top-level folder.
        </DialogDescription>
      </DialogHeader>
      <FolderSelector
        folders={folders}
        isLoading={false}
        value={destination}
        onChangeValue={(value) => setDestination(value.id ? value : TOP_LEVEL)}
        placeholder="Top level"
        excludeFolderId={folder.id}
        filingTargetsOnly
      />
      {moveProblem && (
        <p className="text-muted-foreground text-sm">{moveProblem}</p>
      )}
      <FolderRuleCountNotice ruleCount={ruleCount} change="move" />
      <DialogFooter>
        <Button type="button" variant="outline" onClick={onDone}>
          Cancel
        </Button>
        <Button
          type="submit"
          loading={isMoving}
          disabled={!!moveProblem || ruleCount === "loading"}
        >
          Move
        </Button>
      </DialogFooter>
    </form>
  );
}
