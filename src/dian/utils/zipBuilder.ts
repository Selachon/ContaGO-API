import archiver from "archiver";
import { PassThrough } from "stream";

export interface ZipEntry {
  name: string;
  content: Buffer | string;
}

export function buildZipBuffer(entries: ZipEntry[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const archive = archiver("zip", { zlib: { level: 9 } });
    const output = new PassThrough();
    const chunks: Buffer[] = [];

    output.on("data", (chunk: Buffer) => chunks.push(chunk));
    output.on("end", () => resolve(Buffer.concat(chunks)));
    output.on("error", (error) => reject(error));
    archive.on("error", (error) => reject(error));

    archive.pipe(output);

    for (const entry of entries) {
      archive.append(entry.content, { name: entry.name });
    }

    void archive.finalize();
  });
}
