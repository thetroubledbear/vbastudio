// extension/src/extension.ts
import * as path from "path";
import * as vscode from "vscode";
import { ensureServerPath, getDapServerPath, promptToBrowseForServer } from "./config";
import { configureLaunch, goToProcedure, writeLaunchConfig } from "./configureLaunch";
import { ModuleNode, ModuleTreeProvider, ProcedureNode } from "./moduleTreeProvider";
import { confirmStaleOrProceed, pullModules, pushModules } from "./sync";

// The Modules tree only ever lists Standard modules (ModuleListBuilder.Build filters to
// ModuleKind.Standard - Application.Run can't target Class/Form/Document members), so the
// "Modules" folder and ".bas" extension are always correct here without needing the kind back
// from the server. DapSession's own srcDir convention (<workbookDir>/src/<Kind>/<Name>.ext) is
// the same one SyncEngine's pull writes to.
function moduleSourcePath(workbookPath: string, moduleName: string): string {
  return path.join(path.dirname(workbookPath), "src", "Modules", `${moduleName}.bas`);
}

export function activate(context: vscode.ExtensionContext) {
  const treeProvider = new ModuleTreeProvider();

  const goToProcedureStatusBarItem = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right,
    100
  );
  // Icon and label both say "run", because that is what clicking this does - it writes a
  // launch.json entry for the picked procedure and starts debugging immediately. An earlier
  // $(list-selection) / "Go to Procedure..." pairing read like VS Code's own Go to Symbol, i.e.
  // navigation, which made it easy to run arbitrary VBA against a live workbook by accident.
  goToProcedureStatusBarItem.text = "$(debug-start) VbaStudio";
  goToProcedureStatusBarItem.tooltip = "VbaStudio: Run Procedure...";
  goToProcedureStatusBarItem.command = "vbastudio.goToProcedure";
  goToProcedureStatusBarItem.show();

  context.subscriptions.push(
    goToProcedureStatusBarItem,
    vscode.commands.registerCommand("vbastudio.goToProcedure", () => goToProcedure()),
    vscode.debug.registerDebugAdapterDescriptorFactory("vbastudio", {
      createDebugAdapterDescriptor(): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
        return new vscode.DebugAdapterExecutable(getDapServerPath(), []);
      },
    }),
    vscode.debug.registerDebugConfigurationProvider("vbastudio", {
      async resolveDebugConfiguration(
        _folder: vscode.WorkspaceFolder | undefined,
        config: vscode.DebugConfiguration
      ): Promise<vscode.DebugConfiguration | undefined> {
        if (!(await ensureServerPath("debugging"))) {
          return undefined;
        }
        return config;
      },
    }),
    vscode.commands.registerCommand("vbastudio.setDapServerPath", async () => {
      await promptToBrowseForServer();
      treeProvider.refresh();
    }),
    vscode.commands.registerCommand("vbastudio.configureLaunch", () => configureLaunch()),
    vscode.window.registerTreeDataProvider("vbastudioModules", treeProvider),
    vscode.commands.registerCommand("vbastudio.refreshModules", () => treeProvider.refresh()),
    vscode.commands.registerCommand("vbastudio.pullModules", async () => {
      const succeeded = await pullModules();
      if (succeeded) {
        treeProvider.refresh();
      }
    }),
    vscode.commands.registerCommand("vbastudio.pushModules", () => pushModules()),
    vscode.commands.registerCommand(
      "vbastudio.openModuleSource",
      async (node: ModuleNode | ProcedureNode) => {
        const moduleName = node.kind === "module" ? node.name : node.moduleName;
        const filePath = moduleSourcePath(node.workbookPath, moduleName);
        try {
          const document = await vscode.workspace.openTextDocument(filePath);
          await vscode.window.showTextDocument(document, { preview: false });
        } catch (err: any) {
          vscode.window.showErrorMessage(
            `Could not open "${filePath}". Pull modules first (VbaStudio: Pull Modules from Excel) if it hasn't been synced to disk yet.`
          );
        }
      }
    ),
    vscode.commands.registerCommand(
      "vbastudio.runProcedureFromTree",
      async (node: ProcedureNode) => {
        const folder = vscode.workspace.workspaceFolders?.[0];
        if (!folder) {
          vscode.window.showErrorMessage("Open a folder to configure launch.json.");
          return;
        }

        const shouldProceed = await confirmStaleOrProceed(node.moduleName);
        if (!shouldProceed) {
          return;
        }

        try {
          const config = await writeLaunchConfig(
            folder,
            node.workbookPath,
            node.moduleName,
            node.name
          );
          await vscode.debug.startDebugging(folder, config);
        } catch (err: any) {
          vscode.window.showErrorMessage(err?.message ?? String(err));
        }
      }
    )
  );
}

export function deactivate() {}
