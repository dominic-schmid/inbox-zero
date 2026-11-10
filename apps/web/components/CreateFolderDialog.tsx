"use client";

import { type FormEvent, useState } from "react";
import { toast } from "sonner";
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
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { createFolderAction } from "@/utils/actions/mail";
import { getActionErrorMessage } from "@/utils/error";
import {
  FOLDER_SEPARATOR,
  FOLDER_SEPARATOR_CHARACTER,
  type OutlookFolder,
} from "@/utils/outlook/folders";

const NO_PARENT = { name: "", id: "" };

type CreatedFolder = { id: string; path: string };

export function CreateFolderDialog({
  open,
  onOpenChange,
  emailAccountId,
  folders,
  parent,
  onCreated,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  emailAccountId: string;
  folders: OutlookFolder[];
  /** Preselected parent; the user can still change it. */
  parent?: { id: string; path: string };
  onCreated: (folder: CreatedFolder) => void;
}) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        {/* Remounting per open resets the form to the given parent. */}
        {open && (
          <CreateFolderForm
            emailAccountId={emailAccountId}
            folders={folders}
            parent={parent}
            onCreated={(folder) => {
              onCreated(folder);
              onOpenChange(false);
            }}
            onCancel={() => onOpenChange(false)}
          />
        )}
      </DialogContent>
    </Dialog>
  );
}

function CreateFolderForm({
  emailAccountId,
  folders,
  parent,
  onCreated,
  onCancel,
}: {
  emailAccountId: string;
  folders: OutlookFolder[];
  parent?: { id: string; path: string };
  onCreated: (folder: CreatedFolder) => void;
  onCancel: () => void;
}) {
  const [name, setName] = useState("");
  const [parentValue, setParentValue] = useState(
    parent ? { name: parent.path, id: parent.id } : NO_PARENT,
  );
  const [isCreating, setIsCreating] = useState(false);
  const nameProblem = name.includes(FOLDER_SEPARATOR_CHARACTER)
    ? `Folder names can't contain "${FOLDER_SEPARATOR_CHARACTER}"; it separates folder levels.`
    : null;

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    // React bubbles portal events, so this would also submit an enclosing form
    // such as the rule editor's.
    event.stopPropagation();
    const folderName = name.trim();
    if (!folderName || nameProblem) return;

    setIsCreating(true);
    try {
      const result = await createFolderAction(emailAccountId, {
        name: folderName,
        parentId: parentValue.id || undefined,
      });
      if (result?.serverError || result?.validationErrors || !result?.data) {
        toast.error(getActionErrorMessage(result));
        return;
      }
      const { id, displayName } = result.data;
      onCreated({
        id,
        path: parentValue.id
          ? `${parentValue.name}${FOLDER_SEPARATOR}${displayName}`
          : displayName,
      });
    } finally {
      setIsCreating(false);
    }
  };

  return (
    <form onSubmit={submit} className="space-y-4">
      <DialogHeader>
        <DialogTitle>New folder</DialogTitle>
        <DialogDescription>
          The name is used exactly as typed, slashes included.
        </DialogDescription>
      </DialogHeader>
      <div className="space-y-2">
        <Label htmlFor="new-folder-name">Name</Label>
        <Input
          id="new-folder-name"
          value={name}
          onChange={(event) => setName(event.target.value)}
          autoFocus
          maxLength={255}
          aria-invalid={!!nameProblem}
          aria-describedby={nameProblem ? "new-folder-name-problem" : undefined}
        />
        {nameProblem && (
          <p id="new-folder-name-problem" className="text-destructive text-sm">
            {nameProblem}
          </p>
        )}
      </div>
      <div className="space-y-2">
        <Label>Inside</Label>
        <FolderSelector
          folders={folders}
          isLoading={false}
          value={parentValue}
          onChangeValue={(value) =>
            setParentValue(value.id ? value : NO_PARENT)
          }
          placeholder="Top level"
          filingTargetsOnly
        />
      </div>
      <DialogFooter>
        <Button type="button" variant="outline" onClick={onCancel}>
          Cancel
        </Button>
        <Button
          type="submit"
          loading={isCreating}
          disabled={!name.trim() || !!nameProblem}
        >
          Create
        </Button>
      </DialogFooter>
    </form>
  );
}
