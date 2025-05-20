const { DownloadSingle } = UniverIcons;

const ImportOperationId = 'custom-menu.operation.import';
const ImportButtonOperation = (
  config
) => ({
  id: ImportOperationId,
  type: CommandType.OPERATION,
  handler: async (_accessor) => {
    config?.before?.();
    const univer = _accessor.get(IUniverInstanceService);
    const unitId = univer
      .getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET)
      ?.getUnitId();
    try {
      waitUserSelectExcelFile({
        accept: '.xlsx',
        onSelect: (file) => {
          LuckyExcel.transformExcelToUniver(
            file,
            async (exportJson) => {
              if (unitId) {
                univer.disposeUnit(unitId);
              }
              console.log(exportJson);
              setTimeout(() => {
                const workbook = univer.createUnit(
                  UniverInstanceType.UNIVER_SHEET,
                  exportJson || {}
                );
                config?.after?.({ workbook });
              }, 200);
            },
            (error) => {
              console.log(error);
            }
          );
        },
        onCancel: () => config?.after?.({ error: { message: '取消导入' } }),
        onError: (error) => config?.after?.({ error }),
      });
    } catch (error) {
      config?.after?.({ error });
    }

    return true;
  },
});

const ImportShortcutItem = {
  id: ImportOperationId,
  description: 'shortcut.import',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT | KeyCode.I,
};

function CustomMenuItemImportButtonFactory() {
  return {
    // Bind the command id, clicking the button will trigger this command
    id: ImportOperationId,
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

const CustomImportMenu = (config) => ({
  operation: ImportButtonOperation(config),
  shortcut: ImportShortcutItem,
  menu: CustomMenuItemImportButtonFactory,
  icon: { name: 'DownloadSingle', component: DownloadSingle },
});

const waitUserSelectExcelFile = (params) => {
  const { onSelect, onCancel, onError, accept = '.csv' } = params;
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = accept;
  input.click();
  input.oncancel = () => {
    onCancel?.();
  };
  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    onSelect?.(file);
  };
};
