import fs from "fs";
import XLSX from "xlsx";

export function process_RS(stream: fs.ReadStream, cb: (workbook: XLSX.WorkBook) => void): void {
  const buffers: Buffer[] = [];
  stream.on("data", function (data: Buffer) {
    buffers.push(data);
  });
  stream.on("end", function () {
    const buffer = Buffer.concat(buffers);
    const workbook = XLSX.read(buffer, { type: "buffer" });
    cb(workbook);
  });
}
