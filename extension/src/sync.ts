// extension/src/sync.ts
import * as vscode from "vscode";
import { ensureServerPath, getDapServerPath, isValidServerPath, runServerMode } from "./config";

export async function pullModules(): Promise<boolean> {
  const serverPath = await ensureServerPath("syncing");
  if (!serverPath) {
    return false;
  }

  try {
    // `pull` reports the directory it actually wrote to. That directory is derived from the
    // WORKBOOK's location, which is not necessarily the folder open in VS Code, so it is shown
    // back to the user rather than left implicit.
    const stdout = await runServerMode(serverPath, ["pull"], "pull modules");
    reportSyncResult("pull", parseSyncResult(stdout));
    return true;
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
    return false;
  }
}

interface SyncResultPayload {
  srcDir?: string;
  written: string[];
  deleted: string[];
  conflicts: string[];
}

function parseSyncResult(stdout: string): SyncResultPayload {
  try {
    const parsed = JSON.parse(stdout);
    return {
      srcDir: typeof parsed?.srcDir === "string" ? parsed.srcDir : undefined,
      written: Array.isArray(parsed?.written) ? parsed.written : [],
      deleted: Array.isArray(parsed?.deleted) ? parsed.deleted : [],
      conflicts: Array.isArray(parsed?.conflicts) ? parsed.conflicts : [],
    };
  } catch {
    return { written: [], deleted: [], conflicts: [] };
  }
}

// Conflicts get their own warning dialog (not folded into the info message) because they are
// the one outcome that needs the user to actually do something - both sides changed since the
// last sync, so SyncEngine left that module untouched on both ends rather than guess which side
// wins. Everything else (counts of what moved) is a low-attention info toast.
function reportSyncResult(direction: "pull" | "push", result: SyncResultPayload): void {
  const counts: string[] = [];
  if (result.written.length > 0) {
    counts.push(`${result.written.length} written`);
  }
  if (result.deleted.length > 0) {
    counts.push(`${result.deleted.length} deleted`);
  }
  const countsSuffix = counts.length > 0 ? ` (${counts.join(", ")})` : "";

  const verb = direction === "pull" ? "pulled from Excel" : "pushed to Excel";
  const location = direction === "pull" && result.srcDir ? ` to ${result.srcDir}` : "";
  vscode.window.showInformationMessage(`Modules ${verb}${location}.${countsSuffix}`);

  if (result.conflicts.length > 0) {
    const noun = result.conflicts.length === 1 ? "module" : "modules";
    vscode.window.showWarningMessage(
      `Sync conflict on ${noun}: ${result.conflicts.join(", ")}. ` +
        "Both disk and Excel changed since the last sync - left untouched. Resolve manually, then sync again."
    );
  }
}

export async function pushModules(): Promise<boolean> {
  const serverPath = await ensureServerPath("syncing");
  if (!serverPath) {
    return false;
  }

  try {
    const stdout = await runServerMode(serverPath, ["push"], "push modules");
    reportSyncResult("push", parseSyncResult(stdout));
    return true;
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
    return false;
  }
}

// Fails open: if the check itself can't complete (invalid server path, execFile failure),
// returns false rather than blocking a run over a diagnostic that couldn't run. Deliberately
// uses the raw path check rather than ensureServerPath - this runs as a silent pre-flight, and
// must never pop an error dialog of its own.
export async function isStale(moduleName: string): Promise<boolean> {
  const serverPath = getDapServerPath();
  if (!isValidServerPath(serverPath)) {
    return false;
  }

  try {
    const stdout = await runServerMode(serverPath, ["stale", moduleName], "check module staleness");
    const result = JSON.parse(stdout);
    return !!result.stale;
  } catch {
    return false;
  }
}

// Returns whether the caller should proceed with the run. Not stale -> proceed immediately.
// Stale -> ask; "Push and run" pushes first and only proceeds if the push itself succeeded;
// "Run anyway" proceeds without pushing; anything else (Escape, closing the dialog) cancels.
export async function confirmStaleOrProceed(moduleName: string): Promise<boolean> {
  const stale = await isStale(moduleName);
  if (!stale) {
    return true;
  }

  const action = await vscode.window.showWarningMessage(
    `Module "${moduleName}" has unpushed or unpulled changes.`,
    "Push and run",
    "Run anyway"
  );

  if (action === "Push and run") {
    return await pushModules();
  }

  return action === "Run anyway";
}
