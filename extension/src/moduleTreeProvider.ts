// extension/src/moduleTreeProvider.ts
import * as vscode from "vscode";
import { getDapServerPath, isValidServerPath } from "./config";
import { ModuleListResult, runList } from "./configureLaunch";

export interface MessageNode {
  kind: "message";
  text: string;
  command?: vscode.Command;
}

export interface WorkbookNode {
  kind: "workbook";
  workbookPath: string;
}

export interface ModuleNode {
  kind: "module";
  workbookPath: string;
  name: string;
  procedures: string[];
}

export interface ProcedureNode {
  kind: "procedure";
  workbookPath: string;
  moduleName: string;
  name: string;
}

export type TreeNode = MessageNode | WorkbookNode | ModuleNode | ProcedureNode;

export class ModuleTreeProvider implements vscode.TreeDataProvider<TreeNode> {
  private readonly _onDidChangeTreeData = new vscode.EventEmitter<void>();
  readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

  private cachedResult: ModuleListResult | undefined;

  refresh(): void {
    this.cachedResult = undefined;
    this._onDidChangeTreeData.fire();
  }

  getTreeItem(element: TreeNode): vscode.TreeItem {
    switch (element.kind) {
      case "message": {
        const item = new vscode.TreeItem(element.text, vscode.TreeItemCollapsibleState.None);
        item.command = element.command;
        return item;
      }
      case "workbook": {
        const fileName = element.workbookPath.split(/[\\/]/).pop() ?? element.workbookPath;
        const item = new vscode.TreeItem(fileName, vscode.TreeItemCollapsibleState.Expanded);
        item.tooltip = element.workbookPath;
        return item;
      }
      case "module": {
        const item = new vscode.TreeItem(element.name, vscode.TreeItemCollapsibleState.Collapsed);
        item.description = `${element.procedures.length} procedure${element.procedures.length === 1 ? "" : "s"}`;
        return item;
      }
      case "procedure": {
        const item = new vscode.TreeItem(element.name, vscode.TreeItemCollapsibleState.None);
        item.iconPath = new vscode.ThemeIcon("debug-start");
        item.command = {
          command: "vbastudio.runProcedureFromTree",
          title: "Run/Debug",
          arguments: [element],
        };
        return item;
      }
    }
  }

  async getChildren(element?: TreeNode): Promise<TreeNode[]> {
    if (!element) {
      const serverPath = getDapServerPath();
      if (!isValidServerPath(serverPath)) {
        return [
          {
            kind: "message",
            text: "Set VbaStudio.DapServer.exe path...",
            command: { command: "vbastudio.setDapServerPath", title: "Set path" },
          },
        ];
      }

      try {
        this.cachedResult = await runList(serverPath);
      } catch (err: any) {
        return [{ kind: "message", text: err?.message ?? String(err) }];
      }

      if (!this.cachedResult.modules || this.cachedResult.modules.length === 0) {
        return [{ kind: "message", text: "No modules found in the active workbook." }];
      }

      return [{ kind: "workbook", workbookPath: this.cachedResult.workbookPath }];
    }

    if (element.kind === "workbook") {
      return (this.cachedResult?.modules ?? []).map((m) => ({
        kind: "module" as const,
        workbookPath: element.workbookPath,
        name: m.name,
        procedures: m.procedures,
      }));
    }

    if (element.kind === "module") {
      if (element.procedures.length === 0) {
        return [{ kind: "message", text: `Module "${element.name}" has no procedures.` }];
      }
      return element.procedures.map((p) => ({
        kind: "procedure" as const,
        workbookPath: element.workbookPath,
        moduleName: element.name,
        name: p,
      }));
    }

    return [];
  }
}
