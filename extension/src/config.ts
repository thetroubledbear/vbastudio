// extension/src/config.ts
import { execFile } from "child_process";
import * as fs from "fs";
import { promisify } from "util";
import * as vscode from "vscode";

const execFileAsync = promisify(execFile);

export function getDapServerPath(): string {
  return vscode.workspace.getConfiguration("vbastudio").get<string>("dapServerPath", "");
}

export function isValidServerPath(serverPath: string): boolean {
  return !!serverPath && fs.existsSync(serverPath);
}

export async function promptToBrowseForServer(): Promise<void> {
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    filters: { Executable: ["exe"] },
    openLabel: "Select VbaStudio.DapServer.exe",
  });
  if (!picked || picked.length === 0) {
    return;
  }

  await vscode.workspace
    .getConfiguration("vbastudio")
    .update("dapServerPath", picked[0].fsPath, vscode.ConfigurationTarget.Global);
}

// The single "you haven't configured the server yet" gate for every command that shells out to
// VbaStudio.DapServer.exe. Lives here rather than in any one caller because it used to exist as
// four inline copies whose wording had drifted apart; only the trailing clause legitimately
// varies, so that is the parameter. Returns the validated path, or undefined after having already
// shown the error (and optionally the browse dialog) - callers just bail on undefined.
export async function ensureServerPath(
  actionDescription: string
): Promise<string | undefined> {
  const serverPath = getDapServerPath();
  if (isValidServerPath(serverPath)) {
    return serverPath;
  }

  const action = await vscode.window.showErrorMessage(
    `vbastudio.dapServerPath is not set to a valid VbaStudio.DapServer.exe. Set it before ${actionDescription}.`,
    "Browse..."
  );
  if (action === "Browse...") {
    await promptToBrowseForServer();
  }
  return undefined;
}

// Runs one of the server's one-shot CLI modes (list/pull/push/stale) and returns its stdout.
// Every one of those modes reports failure the same way - one line on stderr, non-zero exit - so
// the "prefer stderr, fall back to the exec error" unwrapping is shared here too. Throws a plain
// Error carrying a ready-to-display message; callers decide whether to surface it in a dialog, a
// tree node, or swallow it.
export async function runServerMode(
  serverPath: string,
  args: string[],
  failureDescription: string
): Promise<string> {
  try {
    const { stdout } = await execFileAsync(serverPath, args);
    return stdout;
  } catch (err: any) {
    const stderrText: string | undefined = err?.stderr?.toString().trim();
    throw new Error(
      stderrText || `Failed to ${failureDescription}: ${err?.message ?? String(err)}`
    );
  }
}
