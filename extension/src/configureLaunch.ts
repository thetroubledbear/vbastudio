// extension/src/configureLaunch.ts
import * as vscode from "vscode";
import { ensureServerPath, runServerMode } from "./config";
import { confirmStaleOrProceed } from "./sync";

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
  const stdout = await runServerMode(serverPath, ["list"], "list modules");
  try {
    return JSON.parse(stdout);
  } catch (err: any) {
    throw new Error(`Failed to list modules: ${err?.message ?? String(err)}`);
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

  const serverPath = await ensureServerPath("configuring a launch target");
  if (!serverPath) {
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

  const shouldProceed = await confirmStaleOrProceed(modulePick.module.name);
  if (!shouldProceed) {
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

// One flat, searchable QuickPick over every runnable procedure across every module - unlike
// configureLaunch()'s two-step module-then-procedure picker, this is the fast path: type a few
// letters of any procedure name (or module name, via matchOnDescription) and run it in one step,
// no confirmation - matching the tree's one-click feel.
export async function goToProcedure(): Promise<void> {
  const folder = vscode.workspace.workspaceFolders?.[0];
  if (!folder) {
    vscode.window.showErrorMessage("Open a folder to configure launch.json.");
    return;
  }

  const serverPath = await ensureServerPath("running a procedure");
  if (!serverPath) {
    return;
  }

  let result: ModuleListResult;
  try {
    result = await runList(serverPath);
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
    return;
  }

  const items = result.modules.flatMap((m) =>
    m.procedures.map((p) => ({
      label: p,
      description: m.name,
      moduleName: m.name,
      procedureName: p,
    }))
  );

  if (items.length === 0) {
    vscode.window.showInformationMessage("No runnable procedures found in the active workbook.");
    return;
  }

  const pick = await vscode.window.showQuickPick(items, {
    placeHolder: "Go to procedure...",
    matchOnDescription: true,
  });
  if (!pick) {
    return;
  }

  const shouldProceed = await confirmStaleOrProceed(pick.moduleName);
  if (!shouldProceed) {
    return;
  }

  try {
    const config = await writeLaunchConfig(
      folder,
      result.workbookPath,
      pick.moduleName,
      pick.procedureName
    );
    await vscode.debug.startDebugging(folder, config);
  } catch (err: any) {
    vscode.window.showErrorMessage(err?.message ?? String(err));
  }
}
