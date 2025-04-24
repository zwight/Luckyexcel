import {
  Disposable,
  ICommandService,
  Inject,
  IConfigService,
} from '@univerjs/core';
import {
  ComponentManager,
  IShortcutService,
  RibbonStartGroup,
  IMenuManagerService,
} from '@univerjs/ui';
import { CUSTOM_PLUGIN_CONFIG, ICustomMenuPluginConfig } from '..';
export class CustomMenuController extends Disposable {
  constructor(
    @ICommandService private readonly _commandService: ICommandService,
    @IMenuManagerService
    private readonly _menuMangerService: IMenuManagerService,
    @Inject(ComponentManager)
    private readonly _componentManager: ComponentManager,
    @IShortcutService private readonly _shortcutService: IShortcutService,
    @IConfigService private readonly _configService: IConfigService
  ) {
    super();

    this._init();
  }

  private _init(): void {
    const config: ICustomMenuPluginConfig =
      this._configService.getConfig(CUSTOM_PLUGIN_CONFIG)!;
    const { menu } = config || {};
    menu.forEach((item, index) => {
      if (item.operation)
        this.disposeWithMe(
          this._commandService.registerCommand(item.operation)
        ); // register command
      if (item.shortcut)
        this.disposeWithMe(
          this._shortcutService.registerShortcut(item.shortcut)
        ); // register shortcut
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
