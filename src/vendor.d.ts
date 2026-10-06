declare module "pako" {
  export interface InflateOptions {
    readonly chunkSize?: number;
    readonly windowBits?: number;
    readonly raw?: boolean;
  }

  export class Inflate {
    constructor(options?: InflateOptions);
    err: number;
    msg: string;
    onData(chunk: Uint8Array): void;
    push(data: Uint8Array, final: boolean): boolean;
  }

  const pako: { readonly Inflate: typeof Inflate };
  export default pako;
}
