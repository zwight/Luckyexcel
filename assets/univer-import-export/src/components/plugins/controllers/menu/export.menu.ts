import {
  type ICommand,
  type IAccessor,
  CommandType,
  IUniverInstanceService,
  UniverInstanceType,
} from '@univerjs/core';
import {
  MenuItemType,
  IShortcutItem,
  KeyCode,
  MetaKeys,
  IMenuButtonItem,
} from '@univerjs/ui';
import { ExportSingle } from '@univerjs/icons';
import { ICustomMenuPulginParams } from '../..';
import LuckyExcel from '@zwight/luckyexcel';

let exportWorker: Worker;
const OperationId = 'custom-menu.operation.export';
const ExportButtonOperation: (config?: ICustomMenuPulginParams) => ICommand = (
  config
) => ({
  id: OperationId,
  type: CommandType.OPERATION,
  handler: async (_accessor: IAccessor) => {
    console.log('Export button operation', config);
    const before = await config?.before?.();

    if (config?.before && !before) return false;
    const univer = _accessor.get(IUniverInstanceService);
    const snapshot = univer
      .getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET)
      ?.getSnapshot();

    const postParams = {
      snapshot: config?.snapshot?.() || snapshot,
      fileName: config?.fileName,
      ...(before || {}),
    };

    LuckyExcel.transformUniverToExcel({
      ...postParams,
      getBuffer: false,
      success: (buffer: Buffer) => {
        console.log('success', buffer);
      },
      error: (error: Error) => {
        console.log('error', error);
        // config?.after?.(error);
        self.postMessage({ error });
      },
    });
    // 导出Worker
    // if (!exportWorker) exportWorker = new ExportWorker();
    // exportWorker.onmessage = (e) => {
    //     if (!e.data) return;
    //     console.log('Received message from worker:', e.data);

    //     if (e.data?.error) config?.after?.(e.data?.error);
    //     if (e.data?.buffer) {
    //         if (before.type && Number(before.type) === FILE_TYPE.csv) {
    //             // 如果是csv可能存在多个sheet，具体需要看返回的结果中有几个sheet
    //             if (isObject(e.data.buffer)) {
    //                 for (const key in e.data.buffer) {
    //                     const element = e.data.buffer[key];
    //                     downloadFile(`${postParams?.fileName}_${key}`, element);
    //                 }
    //             } else {
    //                 downloadFile(postParams?.fileName, e.data.buffer);
    //             }
    //         } else {
    //             downloadFile(postParams?.fileName, e.data.buffer);
    //         }
    //         config?.after?.();
    //     }
    // };
    // exportWorker.postMessage(postParams);
    return true;
  },
});

const ExportShortcutItem: IShortcutItem = {
  id: OperationId,
  description: 'shortcut.export',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT | KeyCode.E,
};

function CustomMenuItemExportButtonFactory(): IMenuButtonItem<string> {
  return {
    // Bind the command id, clicking the button will trigger this command
    id: OperationId,
    // The type of the menu item, in this case, it is a button
    type: MenuItemType.BUTTON,
    // The icon of the button, which needs to be registered in ComponentManager
    icon: 'ExportSingle',
    // The tooltip of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    tooltip: 'customMenu.export',
    // The title of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    title: 'customMenu.export',
  };
}

const CustomExportMenu = (config?: ICustomMenuPulginParams) => ({
  operation: ExportButtonOperation(config),
  shortcut: ExportShortcutItem,
  menu: CustomMenuItemExportButtonFactory,
  icon: { name: 'ExportSingle', component: ExportSingle },
});
const downloadFile = (fileName: string, buffer: Buffer) => {
  const link = document.createElement('a');

  let blob: Blob;
  if (typeof buffer === 'string') {
    blob = new Blob([buffer], { type: 'text/csv;charset=utf-8;' });
  } else {
    blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    });
  }

  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();

  link.addEventListener('click', () => {
    link.remove();
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 200);
  });
};

export default CustomExportMenu;
