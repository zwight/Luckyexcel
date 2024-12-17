interface Window {
    luckysheet: {
        destroy:any,
        create:any,
        getWorker: (label: string) => Worker
    }
}

import '@zwight/exceljs';

declare module '@zwight/exceljs' {
  interface DataValidations {
    add(address: string, validation: DataValidation): void;

    find(address: string): DataValidation | undefined;

    remove(address: string): void;
  }

  interface Worksheet {
    dataValidations: DataValidations;
  }
}