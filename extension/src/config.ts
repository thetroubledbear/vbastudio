// extension/src/config.ts
import * as fs from "fs";
import * as vscode from "vscode";

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
