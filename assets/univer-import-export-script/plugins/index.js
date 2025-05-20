const { Plugin: UniverPlugin, Injector, LocaleService } = UniverCore

const SHEET_CUSTOM_MENU_PLUGIN = 'SHEET_CUSTOM_MENU_PLUGIN';

const AllCustomController = [CustomMenuController];
const enUS = {
    customMenu: {
        export: 'Export',
        import: 'Import',
        save: 'Save',
    },
    shortcut: {
        export: 'Export',
        import: 'Import',
        save: 'Save',
    },
}

const zhCN = {
    customMenu: {
        export: '导出',
        import: '导入',
        save: '保存',
    },
    shortcut: {
        export: '导出',
        import: '导入',
        save: '保存',
    },
};

class UniverSheetsCustomMenuPlugin extends UniverPlugin {
    static type = UniverInstanceType.UNIVER_SHEET;
    static pluginName = SHEET_CUSTOM_MENU_PLUGIN;

    constructor(config, _injector, _localeService, _configService) {
        super();
        
        this._localeService = _localeService;
        this._injector = _injector;
        this._configService = _configService;

        this._configService.setConfig(CUSTOM_PLUGIN_CONFIG, config);
        this._localeService.load({
            zhCN,
            enUS,
        });
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


setDependencies(UniverSheetsCustomMenuPlugin, [Injector, LocaleService, IConfigService], 1)