/** Rules that file into the folder or its subfolders, once known. */
export type FolderRuleCount = number | "loading" | "unknown";

export function FolderRuleCountNotice({
  ruleCount,
  change,
}: {
  ruleCount: FolderRuleCount;
  /** What the folder change does to the paths of rules filing into it. */
  change: "move" | "rename";
}) {
  if (ruleCount === "loading")
    return (
      <p className="text-muted-foreground text-sm">
        Checking which rules file into this folder…
      </p>
    );
  if (ruleCount === 0) return null;

  return (
    <div className="space-y-2 rounded-md border border-amber-300 bg-amber-50 p-3 text-amber-900 text-sm dark:border-amber-700 dark:bg-amber-950 dark:text-amber-200">
      <p>
        {getRuleCountText(ruleCount)} They keep filing into it, and their folder
        path changes to the new {change === "move" ? "location" : "name"}.
      </p>
      <p>
        If those rules are kept in sync from a config file, change the path
        there instead. The next sync sets the old path again, which creates a
        new folder at the old place and splits your mail between the two.
      </p>
    </div>
  );
}

function getRuleCountText(ruleCount: number | "unknown") {
  if (ruleCount === "unknown")
    return "Couldn't check which rules file into this folder, so some may.";
  if (ruleCount === 1)
    return "1 rule files mail into this folder or its subfolders.";
  return `${ruleCount} rules file mail into this folder or its subfolders.`;
}
