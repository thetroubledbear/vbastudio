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
    const srcDir: string | undefined = parseSrcDir(stdout);
    vscode.window.showInformationMessage(
      srcDir ? `Modules pulled from Excel to ${srcDir}.` : "Modules pulled from Excel."
    );
    return true;
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
    return false;
  }
}

function parseSrcDir(stdout: string): string | undefined {
  try {
    const srcDir = JSON.parse(stdout)?.srcDir;
    return typeof srcDir === "string" && srcDir ? srcDir : undefined;
  } catch {
    return undefined;
  }
}

export async function pushModules(): Promise<boolean> {
  const serverPath = await ensureServerPath("syncing");
  if (!serverPath) {
    return false;
  }

  try {
    await runServerMode(serverPath, ["push"], "push modules");
    vscode.window.showInformationMessage("Modules pushed to Excel.");
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
