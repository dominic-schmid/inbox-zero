import type { WheelEvent } from "react";
import { useState } from "react";
import {
  Check,
  ChevronsUpDown,
  FolderIcon,
  ChevronRight,
  Loader2,
  PlusIcon,
  X,
} from "lucide-react";
import { cn } from "@/utils";
import { Button } from "@/components/ui/button";
import {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from "@/components/ui/command";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover";
import {
  FOLDER_SEPARATOR,
  type OutlookFolder,
  withoutNonFilingOutlookFolders,
  withoutOutlookFolder,
} from "@/utils/outlook/folders";
import type { FieldError } from "react-hook-form";
import { CreateFolderDialog } from "@/components/CreateFolderDialog";

interface FolderItemProps {
  displayPath?: string;
  folder: OutlookFolder;
  level: number;
  onSelect: (folderId: string) => void;
  value: { name: string; id: string };
}

function FolderItem({
  folder,
  level,
  value,
  onSelect,
  displayPath,
}: FolderItemProps) {
  return (
    <div key={folder.id}>
      <CommandItem
        key={`${folder.id}-${level}`}
        value={folder.id}
        onSelect={() => onSelect(folder.id)}
        data-folder-id={folder.id}
        data-level={level}
      >
        <Check
          className={cn(
            "mr-2 h-4 w-4",
            value.id === folder.id ? "opacity-100" : "opacity-0",
          )}
        />
        <div className="flex items-center gap-2">
          {level > 0 &&
            Array.from({ length: level }, (_, i) => (
              <ChevronRight key={i} className="h-3 w-3 text-muted-foreground" />
            ))}
          <FolderIcon className="h-4 w-4" />
          <span>{displayPath || folder.displayName}</span>
        </div>
      </CommandItem>
      {folder.childFolders?.map((child) => (
        <div key={child.id} className={""}>
          <FolderItem
            folder={child}
            level={level + 1}
            value={value}
            onSelect={onSelect}
          />
        </div>
      ))}
    </div>
  );
}

interface FolderSelectorProps {
  /** Offers creating a folder from the picker; the new folder gets selected. */
  createFolder?: { emailAccountId: string; onCreated: () => void };
  error?: FieldError;
  /** Leaves this folder and its subfolders out, such as when picking where to move it. */
  excludeFolderId?: string;
  /** Leaves out Drafts, Sent, Deleted Items, Junk and Outbox, with their subfolders. */
  filingTargetsOnly?: boolean;
  folders: OutlookFolder[];
  isLoading: boolean;
  onChangeValue: (value: { name: string; id: string }) => void;
  placeholder?: string;
  value: { name: string; id: string };
}

export function FolderSelector({
  folders: allFolders,
  isLoading,
  value,
  onChangeValue,
  placeholder = "Select a folder...",
  error,
  excludeFolderId,
  createFolder,
  filingTargetsOnly,
}: FolderSelectorProps) {
  const [open, setOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [isCreateOpen, setIsCreateOpen] = useState(false);
  const targets = filingTargetsOnly
    ? withoutNonFilingOutlookFolders(allFolders)
    : allFolders;
  const folders = excludeFolderId
    ? withoutOutlookFolder(targets, excludeFolderId)
    : targets;

  const findFolderById = (
    folderList: OutlookFolder[],
    targetId: string,
  ): OutlookFolder | null => {
    for (const folder of folderList) {
      if (folder.id === targetId) {
        return folder;
      }
      if (folder.childFolders && folder.childFolders.length > 0) {
        const found = findFolderById(folder.childFolders, targetId);
        if (found) return found;
      }
    }
    return null;
  };

  const currentFolderId = value.id;
  const selectedFolder = currentFolderId
    ? findFolderById(folders, currentFolderId)
    : null;

  const filteredFolders =
    searchQuery.trim() === ""
      ? folders.map((folder) => ({ folder, displayPath: folder.displayName }))
      : filterFoldersRecursively(folders, searchQuery.toLowerCase());

  function filterFoldersRecursively(
    folderList: OutlookFolder[],
    query: string,
    parentPath = "",
  ): { folder: OutlookFolder; displayPath: string }[] {
    const results: { folder: OutlookFolder; displayPath: string }[] = [];

    for (const folder of folderList) {
      const currentPath = parentPath
        ? `${parentPath}${FOLDER_SEPARATOR}${folder.displayName}`
        : folder.displayName;
      if (folder.displayName.toLowerCase().includes(query)) {
        results.push({ folder, displayPath: currentPath });
      }
      if (folder.childFolders && folder.childFolders.length > 0) {
        const childResults = filterFoldersRecursively(
          folder.childFolders,
          query,
          currentPath,
        );
        results.push(...childResults);
      }
    }

    return results;
  }

  const buildFolderPath = (folderId: string): string => {
    const folder = findFolderById(folders, folderId);
    if (!folder) return "";

    const findPath = (
      folderList: OutlookFolder[],
      targetId: string,
      currentPath: string[] = [],
    ): string[] | null => {
      for (const f of folderList) {
        const newPath = [...currentPath, f.displayName];

        if (f.id === targetId) {
          return newPath;
        }

        if (f.childFolders && f.childFolders.length > 0) {
          const result = findPath(f.childFolders, targetId, newPath);
          if (result) return result;
        }
      }
      return null;
    };

    const pathParts = findPath(folders, folderId);
    return pathParts ? pathParts.join(FOLDER_SEPARATOR) : folder.displayName;
  };

  const handleFolderSelect = (folderId: string) => {
    const folder = findFolderById(folders, folderId);
    if (folder) {
      const fullPath = buildFolderPath(folderId);
      onChangeValue({
        name: fullPath,
        id: folder.id,
      });
      setOpen(false);
    }
  };

  return (
    <div className="relative">
      <Popover open={open} onOpenChange={setOpen}>
        <PopoverTrigger asChild>
          <Button
            variant="outline"
            role="combobox"
            aria-expanded={open}
            className="w-full justify-between"
            disabled={isLoading}
          >
            <div
              className={cn(
                "flex items-center gap-2 flex-1 min-w-0",
                value.id && !isLoading && "pr-7",
              )}
            >
              {isLoading ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  <span>Loading folders...</span>
                </>
              ) : value.id ? (
                <div className="flex min-w-0 items-center gap-2">
                  <FolderIcon className="h-4 w-4 shrink-0" />
                  <span className="truncate">
                    {value.name || selectedFolder?.displayName || ""}
                  </span>
                </div>
              ) : (
                placeholder
              )}
            </div>
            <ChevronsUpDown className="h-4 w-4 shrink-0 opacity-50" />
          </Button>
        </PopoverTrigger>
        <PopoverContent className="w-[var(--radix-popover-trigger-width)] p-0">
          <Command>
            <CommandInput
              placeholder="Search folders..."
              value={searchQuery}
              onValueChange={setSearchQuery}
            />
            <CommandList
              onWheelCapture={(e: WheelEvent<HTMLDivElement>) => {
                e.preventDefault();
                e.currentTarget.scrollTop += e.deltaY;
              }}
            >
              {isLoading ? (
                <div className="flex items-center justify-center py-6">
                  <Loader2 className="h-4 w-4 animate-spin mr-2" />
                  <span>Loading folders...</span>
                </div>
              ) : (
                <>
                  <CommandEmpty>No folder found.</CommandEmpty>
                  <CommandGroup>
                    {filteredFolders.map(({ folder, displayPath }) => (
                      <FolderItem
                        key={folder.id}
                        folder={folder}
                        level={0}
                        value={value}
                        onSelect={handleFolderSelect}
                        displayPath={displayPath}
                      />
                    ))}
                  </CommandGroup>
                  {createFolder && (
                    <CommandGroup forceMount>
                      <CommandItem
                        forceMount
                        value="create-folder"
                        onSelect={() => {
                          setOpen(false);
                          setIsCreateOpen(true);
                        }}
                      >
                        <PlusIcon className="mr-2 h-4 w-4" />
                        Create folder…
                      </CommandItem>
                    </CommandGroup>
                  )}
                </>
              )}
            </CommandList>
          </Command>
        </PopoverContent>
      </Popover>
      {value.id && !isLoading && (
        <Button
          type="button"
          variant="ghost"
          size="icon2xs"
          className="absolute right-9 top-2 hover:bg-muted"
          onClick={() => onChangeValue({ name: "", id: "" })}
          aria-label="Clear folder selection"
          title="Clear folder selection"
        >
          <X className="h-3 w-3" />
        </Button>
      )}
      {error && (
        <div className="mt-1 text-sm text-red-600 dark:text-red-400">
          {error.message}
        </div>
      )}
      {createFolder && (
        <CreateFolderDialog
          open={isCreateOpen}
          onOpenChange={setIsCreateOpen}
          emailAccountId={createFolder.emailAccountId}
          folders={allFolders}
          onCreated={(folder) => {
            onChangeValue({ name: folder.path, id: folder.id });
            createFolder.onCreated();
          }}
        />
      )}
    </div>
  );
}
