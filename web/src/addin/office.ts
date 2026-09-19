/** Office.js helpers. Safe to import outside Excel — methods throw a clear error. */

interface OfficeFile {
  sliceCount: number;
  getSliceAsync: (
    i: number,
    cb: (result: { status: string; value: { data: unknown }; error?: { message: string } }) => void,
  ) => void;
  closeAsync: () => void;
}

function officeHost(): {
  FileType: { Compressed: string };
  AsyncResultStatus: { Succeeded: string };
  onReady: (cb: () => void) => void;
  context?: {
    document?: {
      url?: string;
      getFileAsync: (
        type: string,
        opts: { sliceSize: number },
        cb: (result: {
          status: string;
          value: OfficeFile;
          error?: { message: string };
        }) => void,
      ) => void;
    };
  };
} | undefined {
  return (globalThis as unknown as { Office?: ReturnType<typeof officeHost> }).Office;
}

export function isInExcel(): boolean {
  return Boolean(officeHost()?.context?.document);
}

export function waitOfficeReady(): Promise<void> {
  return new Promise((resolve) => {
    const o = officeHost();
    if (!o?.onReady) {
      resolve();
      return;
    }
    o.onReady(() => resolve());
  });
}

function sliceToBytes(data: unknown): Uint8Array {
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) {
    const view = data as ArrayBufferView;
    return new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
  }
  if (Array.isArray(data)) return Uint8Array.from(data as number[]);
  throw new Error('Unexpected file slice from Excel');
}

/** Current workbook as .xlsx bytes (in-memory doc, even if unsaved). */
export async function getOpenWorkbookFile(): Promise<File> {
  const Office = officeHost();
  const doc = Office?.context?.document;
  if (!Office || !doc) {
    throw new Error('Open this pane from Excel (Bit add-in), not the browser.');
  }

  const file = await new Promise<OfficeFile>((resolve, reject) => {
    doc.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || 'Could not read the workbook'));
    });
  });

  try {
    const parts: Uint8Array[] = [];
    for (let i = 0; i < file.sliceCount; i++) {
      const slice = await new Promise<{ data: unknown }>((resolve, reject) => {
        file.getSliceAsync(i, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message || `Slice ${i} failed`));
        });
      });
      parts.push(sliceToBytes(slice.data));
    }
    const blob = new Blob(parts as BlobPart[], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const name = (doc.url || 'workbook').replace(/^.*[\\/]/, '') || 'workbook.xlsx';
    const filename = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
    return new File([blob], filename, { type: blob.type });
  } finally {
    file.closeAsync();
  }
}

export async function openWorkbookFromBase64(base64: string): Promise<void> {
  const Excel = (globalThis as unknown as { Excel?: { createWorkbook: (b: string) => void } })
    .Excel;
  if (!Excel?.createWorkbook) {
    throw new Error('This Excel build cannot open a workbook from Bit. Download .xlsx instead.');
  }
  Excel.createWorkbook(base64);
}
