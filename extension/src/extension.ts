// extension/src/extension.ts
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath, promptToBrowseForServer } from "./config";
import { configureLaunch } from "./configureLaunch";

export function activate(context: vscode.ExtensionContext) {
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
    vscode.commands.registerCommand("vbastudio.setDapServerPath", () => promptToBrowseForServer()),
    vscode.commands.registerCommand("vbastudio.configureLaunch", () => configureLaunch())
  );
}

export function deactivate() {}
