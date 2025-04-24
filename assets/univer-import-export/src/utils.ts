export const waitUserSelectExcelFile = (params: {
  onSelect?: (result: File) => void;
  onCancel?: () => void;
  onError?: (error: any) => void;
  accept?: string;
}) => {
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
