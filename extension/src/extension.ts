// extension/src/extension.ts
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath, promptToBrowseForServer } from "./config";
import { configureLaunch, writeLaunchConfig } from "./configureLaunch";
import { ModuleTreeProvider, ProcedureNode } from "./moduleTreeProvider";

export function activate(context: vscode.ExtensionContext) {
  const treeProvider = new ModuleTreeProvider();

  context.subscriptions.push(
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
        const serverPath = getDapServerPath();
        if (!isValidServerPath(serverPath)) {
          const action = await vscode.window.showErrorMessage(
            "vbastudio.dapServerPath is not set to a valid VbaStudio.DapServer.exe. Set it in Settings before debugging.",
            "Browse..."
          );
          if (action === "Browse...") {
            await promptToBrowseForServer();
          }
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
    vscode.commands.registerCommand(
      "vbastudio.runProcedureFromTree",
      async (node: ProcedureNode) => {
        const folder = vscode.workspace.workspaceFolders?.[0];
        if (!folder) {
          vscode.window.showErrorMessage("Open a folder to configure launch.json.");
          return;
        }

        const config = await writeLaunchConfig(
          folder,
          node.workbookPath,
          node.moduleName,
          node.name
        );
        await vscode.debug.startDebugging(folder, config);
      }
    )
  );
}

export function deactivate() {}
