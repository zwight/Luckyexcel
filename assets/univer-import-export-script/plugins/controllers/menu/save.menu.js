
const { SaveSingle } = UniverIcons;
const SaveOperationId = 'custom-menu.operation.save';
const SaveButtonOperation = (
  config
) => ({
  id: SaveOperationId,
  type: CommandType.OPERATION,
  handler: async () => {
    config?.after?.();
    return true;
  },
});

const SaveShortcutItem = {
  id: SaveOperationId,
  description: 'shortcut.save',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | KeyCode.S,
};

function CustomMenuItemSaveButtonFactory(
  accessor
) {
  return {
    // Bind the command id, clicking the button will trigger this command
    id: SaveOperationId,
    // The type of the menu item, in this case, it is a button
    type: MenuItemType.BUTTON,
    // The icon of the button, which needs to be registered in ComponentManager
    icon: 'SaveSingle',
    // The tooltip of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    tooltip: 'customMenu.save',
    // The title of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    title: 'customMenu.save',
  };
}
console.log(SaveSingle)
const CustomSaveMenu = (config) => ({
  id: SaveOperationId,
  operation: SaveButtonOperation(config),
  shortcut: SaveShortcutItem,
  menu: CustomMenuItemSaveButtonFactory,
  icon: { name: 'SaveSingle', component: SaveSingle },
});
