// extension/src/configureLaunch.ts
import { execFile } from "child_process";
import { promisify } from "util";
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath, promptToBrowseForServer } from "./config";

const execFileAsync = promisify(execFile);

export interface ModuleListing {
  name: string;
  procedures: string[];
}

export interface ModuleListResult {
  workbookPath: string;
  modules: ModuleListing[];
}

// Shells out to `<serverPath> list` and parses its JSON stdout. Throws a plain Error with a
// ready-to-display message on failure - callers decide how to surface it (a dialog for a
// deliberately-triggered command, a tree node for a background/render-time fetch).
export async function runList(serverPath: string): Promise<ModuleListResult> {
  try {
    const { stdout } = await execFileAsync(serverPath, ["list"]);
    return JSON.parse(stdout);
  } catch (err: any) {
    const stderrText: string | undefined = err?.stderr?.toString().trim();
    throw new Error(stderrText || `Failed to list modules: ${err?.message ?? String(err)}`);
  }
}

// Writes or updates (matched by "name") a launch.json entry for the given module/procedure.
// Returns the resolved config - the caller decides whether/when to start debugging with it.
export async function writeLaunchConfig(
  folder: vscode.WorkspaceFolder,
  workbookPath: string,
  moduleName: string,
  procedureName: string
): Promise<vscode.DebugConfiguration> {
  const entryPoint = `${moduleName}.${procedureName}`;
  const name = `VbaStudio: ${entryPoint}`;
  const newConfig: vscode.DebugConfiguration = {
    type: "vbastudio",
    request: "launch",
    name,
    program: workbookPath,
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

  return newConfig;
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
    result = await runList(serverPath);
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
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

  const newConfig = await writeLaunchConfig(
    folder,
    result.workbookPath,
    modulePick.module.name,
    procedurePick
  );

  const start = await vscode.window.showInformationMessage("Start debugging now?", "Yes");
  if (start === "Yes") {
    await vscode.debug.startDebugging(folder, newConfig);
  }
}
