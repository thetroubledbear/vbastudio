// extension/src/sync.ts
import { execFile } from "child_process";
import { promisify } from "util";
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath, promptToBrowseForServer } from "./config";

const execFileAsync = promisify(execFile);

async function ensureServerPath(): Promise<string | undefined> {
  const serverPath = getDapServerPath();
  if (isValidServerPath(serverPath)) {
    return serverPath;
  }

  const action = await vscode.window.showErrorMessage(
    "vbastudio.dapServerPath is not set to a valid VbaStudio.DapServer.exe. Set it before syncing.",
    "Browse..."
  );
  if (action === "Browse...") {
    await promptToBrowseForServer();
  }
  return undefined;
}

export async function pullModules(): Promise<boolean> {
  const serverPath = await ensureServerPath();
  if (!serverPath) {
    return false;
  }

  try {
    await execFileAsync(serverPath, ["pull"]);
    vscode.window.showInformationMessage("Modules pulled from Excel.");
    return true;
  } catch (err: any) {
    const stderrText: string | undefined = err?.stderr?.toString().trim();
    vscode.window.showErrorMessage(
      stderrText || `Failed to pull modules: ${err?.message ?? String(err)}`
    );
    return false;
  }
}

export async function pushModules(): Promise<boolean> {
  const serverPath = await ensureServerPath();
  if (!serverPath) {
    return false;
  }

  try {
    await execFileAsync(serverPath, ["push"]);
    vscode.window.showInformationMessage("Modules pushed to Excel.");
    return true;
  } catch (err: any) {
    const stderrText: string | undefined = err?.stderr?.toString().trim();
    vscode.window.showErrorMessage(
      stderrText || `Failed to push modules: ${err?.message ?? String(err)}`
    );
    return false;
  }
}

// Fails open: if the check itself can't complete (invalid server path, execFile failure),
// returns false rather than blocking a run over a diagnostic that couldn't run.
export async function isStale(moduleName: string): Promise<boolean> {
  const serverPath = getDapServerPath();
  if (!isValidServerPath(serverPath)) {
    return false;
  }

  try {
    const { stdout } = await execFileAsync(serverPath, ["stale", moduleName]);
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
