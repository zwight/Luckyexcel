const { Disposable, setDependencies, ICommandService, IConfigService } = UniverCore;
const { RibbonStartGroup, ComponentManager, IMenuManagerService, IShortcutService } = UniverUi;
const CUSTOM_PLUGIN_CONFIG = 'CUSTOM_PLUGIN_CONFIG';
class CustomMenuController extends Disposable {
  constructor(_commandService, _menuMangerService, _componentManager, _shortcutService, _configService) {
    super();

    this._commandService = _commandService;
    this._menuMangerService = _menuMangerService;
    this._componentManager = _componentManager;
    this._shortcutService = _shortcutService;
    this._configService = _configService;

    this._init();
  }
  _init() {
    // console.log(this)
    const config = this._configService.getConfig(CUSTOM_PLUGIN_CONFIG);
    console.log(config)
    const { menu } = config || {};
    if (!menu) return;
    menu.forEach((item, index) => {
      if (item.operation)
        this.disposeWithMe(
          this._commandService.registerCommand(item.operation)
        ); // register command
      if (item.shortcut)
        this.disposeWithMe(
          this._shortcutService.registerShortcut(item.shortcut)
        ); // register shortcut
        console.log(this._componentManager, UniverIcons)
      if (item.icon)
        
        this.disposeWithMe(
          this._componentManager.register(item.icon.name, item.icon.component)
        ); // register icon component

      if (item.menu) {
        this._menuMangerService.mergeMenu({
          [RibbonStartGroup.HISTORY]: {
            [item.menu().id]: {
              order: -((menu.length || 0) - index) - 2,
              menuItemFactory: item.menu,
            },
          },
        });
      }
    });
  }
}

setDependencies(CustomMenuController, [ICommandService, IMenuManagerService, ComponentManager, IShortcutService, IConfigService])
