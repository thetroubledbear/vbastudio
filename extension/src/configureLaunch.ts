// extension/src/configureLaunch.ts
import { execFile } from "child_process";
import { promisify } from "util";
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath, promptToBrowseForServer } from "./config";

const execFileAsync = promisify(execFile);

interface ModuleListing {
  name: string;
  procedures: string[];
}

interface ModuleListResult {
  workbookPath: string;
  modules: ModuleListing[];
}

export async function configureLaunch(): Promise<void> {
  const folder = vscode.workspace.workspaceFolders?.[0];
  if (!folder) {
    vscode.window.showErrorMessage("Open a folder to configure launch.json.");
    return;
  }

  const serverPath = getDapServerPath();
  if (!isValidServerPath(serverPath)) {
    const action = await vscode.window.showErrorMessage(
      "vbastudio.dapServerPath is not set to a valid VbaStudio.DapServer.exe. Set it before configuring a launch target.",
      "Browse..."
    );
    if (action === "Browse...") {
      await promptToBrowseForServer();
    }
    return;
  }

  let result: ModuleListResult;
  try {
    const { stdout } = await execFileAsync(serverPath, ["list"]);
    result = JSON.parse(stdout);
  } catch (err: any) {
    const stderrText: string | undefined = err?.stderr?.toString().trim();
    vscode.window.showErrorMessage(
      stderrText || `Failed to list modules: ${err?.message ?? String(err)}`
    );
    return;
  }

  if (!result.modules || result.modules.length === 0) {
    vscode.window.showInformationMessage("No modules found in the active workbook.");
    return;
  }

  const modulePick = await vscode.window.showQuickPick(
    result.modules.map((m) => ({
      label: m.name,
      description: `${m.procedures.length} procedure${m.procedures.length === 1 ? "" : "s"}`,
      module: m,
    })),
    { placeHolder: "Select a module" }
  );
  if (!modulePick) {
    return;
  }

  if (modulePick.module.procedures.length === 0) {
    vscode.window.showInformationMessage(`Module "${modulePick.module.name}" has no procedures.`);
    return;
  }

  const procedurePick = await vscode.window.showQuickPick(modulePick.module.procedures, {
    placeHolder: "Select a procedure",
  });
  if (!procedurePick) {
    return;
  }

  const entryPoint = `${modulePick.module.name}.${procedurePick}`;
  const name = `VbaStudio: ${entryPoint}`;
  const newConfig = {
    type: "vbastudio",
    request: "launch",
    name,
    program: result.workbookPath,
    entryPoint,
  };

  const launchConfig = vscode.workspace.getConfiguration("launch", folder.uri);
  const configurations = launchConfig.get<any[]>("configurations", []).slice();
  const existingIndex = configurations.findIndex((c) => c.name === name);
  if (existingIndex >= 0) {
    configurations[existingIndex] = newConfig;
  } else {
    configurations.push(newConfig);
  }
  await launchConfig.update(
    "configurations",
    configurations,
    vscode.ConfigurationTarget.WorkspaceFolder
  );

  const start = await vscode.window.showInformationMessage("Start debugging now?", "Yes");
  if (start === "Yes") {
    await vscode.debug.startDebugging(folder, newConfig);
  }
}
