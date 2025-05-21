import {
  LocaleService,
  Plugin,
  Inject,
  Injector,
  UniverInstanceType,
  ICommand,
  Dependency,
  IConfigService,
  Univer,
  IAccessor,
} from '@univerjs/core';
import { IMenuButtonItem, IShortcutItem } from '@univerjs/ui';

import zhCN from './locale/zh-CN';
import enUS from './locale/en-US';
import { CustomMenuController } from './controllers/custom-menu.controller';

const SHEET_CUSTOM_MENU_PLUGIN = 'SHEET_CUSTOM_MENU_PLUGIN';

export interface ICustomMenuPluginConfig {
  instance: Univer;
  menu: Array<{
    operation: ICommand<object, boolean>;
    shortcut?: IShortcutItem<object>;
    menu: () => IMenuButtonItem<string>;
    icon?: { name: string; component: any };
  }>;
}
export interface UniverMenuConfig {
  id: string;
  operation: ICommand<object, boolean>;
  shortcut?: IShortcutItem<object>;
  menu: () => IMenuButtonItem<string>;
  icon?: { name: string; component: any };
  onlyOperation?: boolean;
}
export interface ICustomMenuPulginParams {
  [key: string]: any;
  before?: () => void | Promise<any>;
  after?: (param?: any) => void;
}
export const CUSTOM_PLUGIN_CONFIG = 'CUSTOM_PLUGIN_CONFIG';
const AllCustomController = [CustomMenuController];

export class UniverSheetsCustomMenuPlugin extends Plugin {
  static override type = UniverInstanceType.UNIVER_SHEET;
  static override pluginName = SHEET_CUSTOM_MENU_PLUGIN;

  constructor(
    readonly config: ICustomMenuPluginConfig,
    @Inject(Injector) protected readonly _injector: Injector,
    @Inject(LocaleService) private readonly _localeService: LocaleService,
    @IConfigService private readonly _configService: IConfigService
  ) {
    super();
    this._localeService.load({
      zhCN,
      enUS,
    });
    this._configService.setConfig(CUSTOM_PLUGIN_CONFIG, config);
  }

  onReady(): void {
    AllCustomController.forEach((d) => this._injector.get(d));
  }

  override onStarting(): void {
    ([AllCustomController] as Dependency[]).forEach((d) =>
      this._injector.add(d)
    );
  }
}
