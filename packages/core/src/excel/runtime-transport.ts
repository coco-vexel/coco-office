import {
  deserializeExcelError,
  type ExcelRpcError,
  type ExcelRpcMessage,
  type ExcelRpcRequest,
  type ExcelRpcResponse,
} from './runtime-rpc.js';

export interface ExcelWorkerPortLike {
  postMessage(data: unknown, transferList?: unknown[]): void;
  onMessage(listener: (data: unknown) => void): void;
  onError(listener: (err: unknown) => void): void;
  terminate(): Promise<void>;
}

interface PendingCall {
  resolve: (value: unknown) => void;
  reject: (err: unknown) => void;
  timeoutHandle?: ReturnType<typeof setTimeout>;
}

export class ExcelWorkerTransport {
  private readonly pending = new Map<string, PendingCall>();
  private nextId = 0;
  private disposed = false;

  constructor(private readonly port: ExcelWorkerPortLike) {
    port.onMessage((data) => this.handleIncoming(data));
    port.onError((error) => this.rejectAll(error));
  }

  call<T>(method: string, args: unknown[], timeoutMs = 0): Promise<T> {
    if (this.disposed) throw new Error('ExcelWorkerTransport has been disposed');
    const id = `excel-${Date.now()}-${this.nextId++}`;
    const req: ExcelRpcRequest = { kind: 'request', id, method, args };

    return new Promise<T>((resolve, reject) => {
      let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
      if (timeoutMs > 0) {
        timeoutHandle = setTimeout(() => {
          this.pending.delete(id);
          reject(new Error(`excel worker call timeout after ${timeoutMs}ms: ${method}`));
        }, timeoutMs);
      }

      this.pending.set(id, {
        resolve: (value) => resolve(value as T),
        reject,
        timeoutHandle,
      });
      this.port.postMessage(req);
    });
  }

  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    this.rejectAll(new Error('ExcelWorkerTransport disposed'));
    await this.port.terminate();
  }

  private handleIncoming(data: unknown): void {
    if (data === null || typeof data !== 'object') return;
    const msg = data as ExcelRpcMessage;
    if (msg.kind !== 'response' && msg.kind !== 'error') return;

    const pending = this.pending.get(msg.id);
    if (!pending) return;
    this.pending.delete(msg.id);
    if (pending.timeoutHandle) clearTimeout(pending.timeoutHandle);

    if (msg.kind === 'response') {
      pending.resolve((msg as ExcelRpcResponse).result);
    } else {
      pending.reject(deserializeExcelError((msg as ExcelRpcError).error));
    }
  }

  private rejectAll(error: unknown): void {
    for (const [id, pending] of this.pending) {
      if (pending.timeoutHandle) clearTimeout(pending.timeoutHandle);
      pending.reject(error);
      this.pending.delete(id);
    }
  }
}

export async function createNodeExcelWorkerPort(bundleUrl: string | URL): Promise<ExcelWorkerPortLike> {
  const { Worker } = await import('node:worker_threads');
  const worker = new Worker(bundleUrl);
  return {
    postMessage(data, transferList) {
      worker.postMessage(data, transferList as readonly import('node:worker_threads').TransferListItem[] | undefined);
    },
    onMessage(listener) {
      worker.on('message', listener);
    },
    onError(listener) {
      worker.on('error', listener);
      worker.on('exit', (code) => {
        if (code !== 0) listener(new Error(`excel worker exited with code ${code}`));
      });
    },
    async terminate() {
      await worker.terminate();
    },
  };
}
