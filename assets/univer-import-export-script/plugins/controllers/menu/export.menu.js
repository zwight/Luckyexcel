const { ExportSingle } = UniverIcons;
// import LuckyExcel from '@zwight/luckyexcel';

const OperationId = 'custom-menu.operation.export';
const ExportButtonOperation = (
  config
) => ({
  id: OperationId,
  type: CommandType.OPERATION,
  handler: async (_accessor) => {
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
      success: (buffer) => {
        console.log('success', buffer);
      },
      error: (error) => {
        console.log('error', error);
        // config?.after?.(error);
        self.postMessage({ error });
      },
    });
    return true;
  },
});

const ExportShortcutItem = {
  id: OperationId,
  description: 'shortcut.export',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT | KeyCode.E,
};

function CustomMenuItemExportButtonFactory() {
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

const CustomExportMenu = (config) => ({
  operation: ExportButtonOperation(config),
  shortcut: ExportShortcutItem,
  menu: CustomMenuItemExportButtonFactory,
  icon: { name: 'ExportSingle', component: ExportSingle },
});

export default CustomExportMenu;
