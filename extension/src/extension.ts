import * as fs from "fs";
import * as vscode from "vscode";

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(
    vscode.debug.registerDebugAdapterDescriptorFactory("vbastudio", {
      createDebugAdapterDescriptor(): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
        const serverPath = vscode.workspace
          .getConfiguration("vbastudio")
          .get<string>("dapServerPath", "");
        return new vscode.DebugAdapterExecutable(serverPath, []);
      },
    }),
    vscode.debug.registerDebugConfigurationProvider("vbastudio", {
      resolveDebugConfiguration(
        _folder: vscode.WorkspaceFolder | undefined,
        config: vscode.DebugConfiguration
      ): vscode.ProviderResult<vscode.DebugConfiguration> {
        const serverPath = vscode.workspace
          .getConfiguration("vbastudio")
          .get<string>("dapServerPath", "");
        if (!serverPath || !fs.existsSync(serverPath)) {
          vscode.window.showErrorMessage(
            "vbastudio.dapServerPath is not set to a valid VbaStudio.DapServer.exe. Set it in Settings before debugging."
          );
          return undefined;
        }
        return config;
      },
    })
  );
}

export function deactivate() {}
