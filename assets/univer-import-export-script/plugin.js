const { Plugin, UniverInstanceType, Injector, setDependencies, LocaleService, ConfigService } = UniverCore

import { CustomMenuController } from './controllers/custom-menu.controller';

const SHEET_CUSTOM_MENU_PLUGIN = 'SHEET_CUSTOM_MENU_PLUGIN';

const CUSTOM_PLUGIN_CONFIG = 'CUSTOM_PLUGIN_CONFIG';
const AllCustomController = [CustomMenuController];
const enUS = {
    customMenu: {
        export: 'Export',
        import: 'Import',
    },
    shortcut: {
        export: 'Export',
        import: 'Import',
    },
}

const zhCN = {
    customMenu: {
        export: '导出',
        import: '导入',
    },
    shortcut: {
        export: '导出',
        import: '导入',
    },
};

class UniverSheetsCustomMenuPlugin extends Plugin {
    static type = UniverInstanceType.UNIVER_SHEET;
    static pluginName = SHEET_CUSTOM_MENU_PLUGIN;

    constructor(config, _injector, _localeService, _configService) {
        super();
        this._localeService.load({
            zhCN,
            enUS,
        });
        this._configService.setConfig(CUSTOM_PLUGIN_CONFIG, config);
    }

    onReady() {
        AllCustomController.forEach((d) => this._injector.get(d));
    }

    onStarting() {
        ([AllCustomController]).forEach((d) =>
            this._injector.add(d)
        );
    }
}


setDependencies(UniverSheetsCustomMenuPlugin, [Injector, LocaleService, ConfigService])