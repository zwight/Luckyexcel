<script setup lang="ts">
import { onMounted, ref } from 'vue';
import { LocaleType, Tools, Univer, UniverInstanceType } from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import { defaultTheme } from '@univerjs/design';

import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverUIPlugin } from '@univerjs/ui';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';
import { UniverSheetsPrintPlugin } from '@univerjs-pro/sheets-print';
import { UniverSheetsChartPlugin } from '@univerjs-pro/sheets-chart';
import { UniverSheetsChartUIPlugin } from '@univerjs-pro/sheets-chart-ui';
import { UniverLicensePlugin } from '@univerjs-pro/license';

import { UniverDocsDrawingPlugin } from '@univerjs/docs-drawing';
import { IImageIoService, UniverDrawingPlugin } from '@univerjs/drawing';
import { UniverDrawingUIPlugin } from '@univerjs/drawing-ui';
import { UniverSheetsDrawingPlugin } from '@univerjs/sheets-drawing';
import { UniverSheetsDrawingUIPlugin } from '@univerjs/sheets-drawing-ui';

import DesignZhCN from '@univerjs/design/locale/zh-CN';
import UIZhCN from '@univerjs/ui/locale/zh-CN';
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
import SheetsFormulaUIZhCN from '@univerjs/sheets-formula-ui/locale/zh-CN';
import SheetsNumfmtUIZhCN from '@univerjs/sheets-numfmt-ui/locale/zh-CN';
import SheetsPrintPluginZhCN from '@univerjs-pro/sheets-print/locale/zh-CN';
import SheetsChartZhCN from '@univerjs-pro/sheets-chart/locale/zh-CN';
import SheetsChartUIZhCN from '@univerjs-pro/sheets-chart-ui/locale/zh-CN';
import DrawingUIZhCN from '@univerjs/drawing-ui/locale/zh-CN';
import SheetsDrawingUIZhCN from '@univerjs/sheets-drawing-ui/locale/zh-CN';

// 这里的 Facade API 是可选的，你可以根据自己的需求来决定是否引入
import '@univerjs/sheets/facade';

import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/docs-ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/sheets-formula-ui/lib/index.css';
import '@univerjs/sheets-numfmt-ui/lib/index.css';
import '@univerjs-pro/sheets-print/lib/index.css';
import '@univerjs/drawing-ui/lib/index.css';
import '@univerjs-pro/sheets-chart-ui/lib/index.css';
import '@univerjs/sheets-drawing-ui/lib/index.css';

import { UniverSheetsCustomMenuPlugin } from './plugins';
import CustomImportMenu from './plugins/controllers/menu/import.menu';
import CustomExportMenu from './plugins/controllers/menu/export.menu';
import CustomSaveMenu from './plugins/controllers/menu/save.menu';

const univerAPI = ref<Funiver | null>(null);
const loading = ref(false);

onMounted(() => {
  init();
});
const init = () => {
  const univer = new Univer({
    theme: defaultTheme,
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: Tools.deepMerge(
        DesignZhCN,
        UIZhCN,
        DocsUIZhCN,
        SheetsZhCN,
        SheetsUIZhCN,
        SheetsFormulaUIZhCN,
        SheetsNumfmtUIZhCN,
        SheetsPrintPluginZhCN,
        SheetsChartZhCN,
        SheetsChartUIZhCN,
        DrawingUIZhCN,
        SheetsDrawingUIZhCN
      ),
    },
  });
  univer.registerPlugin(UniverLicensePlugin, {
    license: `1846850823316525117-1-eyJpIjoiMTg0Njg1MDgyMzMxNjUyNTExNyIsInYiOiIxIiwicCI6IkVFNTRrV0NxRkdUOUNWcEJsK2ZNWnBpekgxQ3pDV2dzWEMyKzB1c3RRRFk9IiwiZG0iOlsiKi51bml2ZXIuYWkiLCIqLnVuaXZlci5wbHVzIl0sInJ0IjozLCJmdCI6eyJ1ZiI6eyJtdSI6MjE0NzQ4MzY0NiwiZXQiOjE3Njc4ODc5OTksIm1tIjoyMTQ3NDgzNjQ2LCJjdSI6MjE0NzQ4MzY0Nn0sInNmIjp7ImV0IjoxNzY3ODg3OTk5LCJydiI6dHJ1ZSwicHRuIjoyMTQ3NDgzNjQ2LCJtaXMiOjIxNDc0ODM2NDYsIm1wbiI6MjE0NzQ4MzY0NiwibmMiOjIxNDc0ODM2NDZ9LCJkZiI6eyJldCI6MTc2Nzg4Nzk5OSwicnYiOnRydWUsIm1pcyI6MjE0NzQ4MzY0NiwibXBuIjoyMTQ3NDgzNjQ2fSwid3NmIjp7ImV0IjoxNzY3ODg3OTk5LCJobiI6MjE0NzQ4MzY0Nn19LCJ1ZCI6MTc2Nzg4Nzk5OSwiYXQiOjE3MzYzMzA2ODYsImUiOiJkZXZlbG9wZXJAdW5pdmVyLmFpIiwiZCI6OCwibiI6MTUyfQ==-bjVjDKN3EaHAvf4ySUFxnjHsXiKnk1RSiiOD7RaeXkXRHheSdiMjfctXscfYQxJdleRIA9hSSlc4TO9Ncy/qCQ==-1767887999`
  });

  univerAPI.value = FUniver.newAPI(univer);

  
  univer.registerPlugin(UniverRenderEnginePlugin);
  univer.registerPlugin(UniverFormulaEnginePlugin);

  univer.registerPlugin(UniverUIPlugin, {
    container: 'app',
  });

  univer.registerPlugin(UniverDocsPlugin);
  univer.registerPlugin(UniverDocsUIPlugin);

  univer.registerPlugin(UniverSheetsPlugin);
  univer.registerPlugin(UniverSheetsUIPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin);
  univer.registerPlugin(UniverSheetsFormulaUIPlugin);
  univer.registerPlugin(UniverSheetsNumfmtPlugin);
  univer.registerPlugin(UniverSheetsNumfmtUIPlugin);
  univer.registerPlugin(UniverDrawingPlugin);
  univer.registerPlugin(UniverDrawingUIPlugin);
  univer.registerPlugin(UniverSheetsDrawingPlugin);
  univer.registerPlugin(UniverSheetsDrawingUIPlugin);

  univer.registerPlugin(UniverSheetsPrintPlugin);

  univer.registerPlugin(UniverSheetsChartPlugin);
  univer.registerPlugin(UniverSheetsChartUIPlugin);

  univer.registerPlugin(UniverSheetsCustomMenuPlugin, {
    instance: univer,
    menu: [
      CustomSaveMenu({
        after: () => {
          const saveData = getUniverSnapshot();
          console.log(saveData);
        },
      }),
      CustomImportMenu({
        before: () => {
          loading.value = true;
        },
        after: ({ error }) => {
          loading.value = false;
          if (error) {
            Message.error(error.message || '导入失败');
            return;
          }
          Message.success('导入成功');
        },
      }),
      CustomExportMenu({
        snapshot: getUniverSnapshot,
        after: (res) => {
          loading.value = false;
          if (res) {
            Message.error(res.message || '导出失败');
          } else {
            Message.success('导出成功');
          }
        },
      }),
    ],
  });

  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
};

const getUniverSnapshot = () => {
  const activeWorkbook = univerAPI.value?.getActiveWorkbook();
  if (!activeWorkbook) {
    throw new Error('Workbook is not initialized');
  }
  return activeWorkbook.save();
};
</script>

<template>
  <div id="univer"></div>
</template>

<style lang="less">
#univer {
  width: 100%;
  height: 100%;
}
</style>
