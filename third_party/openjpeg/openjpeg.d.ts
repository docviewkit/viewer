export interface OpenJpegModule {
  _malloc(size: number): number;
  _free(pointer: number): void;
  _jp2_decode(
    pointer: number,
    size: number,
    components: number,
    indexed: boolean,
    alpha: boolean,
    reducePower: number,
  ): number;
  writeArrayToMemory(bytes: Uint8Array, pointer: number): void;
  imageData: Uint8ClampedArray | null;
  errorMessages?: string;
}

export interface OpenJpegOptions {
  locateFile?(name: string): string;
  wasmBinary?: Uint8Array;
  warn?(message: string): void;
}

export default function OpenJPEG(options?: OpenJpegOptions): Promise<OpenJpegModule>;
