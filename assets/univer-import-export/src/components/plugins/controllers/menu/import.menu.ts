import {
  MenuItemType,
  IShortcutItem,
  KeyCode,
  MetaKeys,
  IMenuButtonItem,
} from '@univerjs/ui';
import {
  CommandType,
  IAccessor,
  ICommand,
  IUniverInstanceService,
  IWorkbookData,
  UniverInstanceType,
  Workbook,
} from '@univerjs/core';
import { DownloadSingle } from '@univerjs/icons';

import { ICustomMenuPulginParams } from '../..';
import { waitUserSelectExcelFile } from '../../../../utils';
import { UnitModel } from '@univerjs/core/lib/types/common/unit';
import LuckyExcel from '@zwight/luckyexcel';

const OperationId = 'custom-menu.operation.import';
const ImportButtonOperation: (config?: ICustomMenuPulginParams) => ICommand = (
  config
) => ({
  id: OperationId,
  type: CommandType.OPERATION,
  handler: async (_accessor: IAccessor) => {
    config?.before?.();
    const univer = _accessor.get(IUniverInstanceService);
    const unitId = univer
      .getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET)
      ?.getUnitId();
    try {
      waitUserSelectExcelFile({
        accept: '.xlsx',
        onSelect: (file: File) => {
          LuckyExcel.transformExcelToUniver(
            file,
            async (exportJson: any) => {
              if (unitId) {
                univer.disposeUnit(unitId);
              }
              console.log(exportJson);
              setTimeout(() => {
                const workbook = univer.createUnit<IWorkbookData, Workbook>(
                  UniverInstanceType.UNIVER_SHEET,
                  exportJson || {}
                );
                config?.after?.({ workbook });
              }, 200);
            },
            (error: any) => {
              console.log(error);
            }
          );
        },
        onCancel: () => config?.after?.({ error: { message: '取消导入' } }),
        onError: (error: any) => config?.after?.({ error }),
      });
    } catch (error) {
      config?.after?.({ error });
    }

    return true;
  },
});

const ImportShortcutItem: IShortcutItem = {
  id: OperationId,
  description: 'shortcut.import',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT | KeyCode.I,
};

function CustomMenuItemImportButtonFactory(): IMenuButtonItem<string> {
  return {
    // Bind the command id, clicking the button will trigger this command
    id: OperationId,
    // The type of the menu item, in this case, it is a button
    type: MenuItemType.BUTTON,
    // The icon of the button, which needs to be registered in ComponentManager
    icon: 'DownloadSingle',
    // The tooltip of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    tooltip: 'customMenu.import',
    // The title of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    title: 'customMenu.import',
    // The button position can be configured in the toolbar or context menu using MenuPosition. If it is a sheet, you can also use SheetMenuPosition to configure the row header, column header, or sheet bar context menu
    // positions: [MenuPosition.TOOLBAR_START],
    // group: MenuGroup.TOOLBAR_HISTORY,
  };
}

const CustomImportMenu = (config?: ICustomMenuPulginParams) => ({
  operation: ImportButtonOperation(config),
  shortcut: ImportShortcutItem,
  menu: CustomMenuItemImportButtonFactory,
  icon: { name: 'DownloadSingle', component: DownloadSingle },
});

export default CustomImportMenu;
