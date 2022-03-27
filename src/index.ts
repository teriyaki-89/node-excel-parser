import XLSX from "xlsx";
import fs from "fs";
import { process_RS } from "./readable-stream";

function prepareXlsxIterator(workbook: XLSX.WorkBook, filename: string, sheetName: string) {
  const columnsObject = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 0 })[0] as Record<
    string,
    unknown
  >;
  const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]) as Record<string, any>[];
  const batches: Record<string, any>[] = [];

  for (const excelRecord of worksheet) {
    const bulkRecord: Record<string, any> = {};
    Object.keys(columnsObject).map((key: string) => {
      let fieldName = columnsObject[key] as string;
      fieldName = prepareDatabaseFieldName(fieldName);
      bulkRecord[fieldName] = excelRecord[key] ? (excelRecord[key] as Record<string, any>) : null;
    });
    batches.push(bulkRecord);
  }

  console.log(batches);
  return batches;
}

function prepareDatabaseFieldName(excelField: string) {
  return excelField.replace(/\s/g, "-");
}

try {
  const filename = "NEM Generation Information Feb 2022.xlsx";
  const sheetName = "Scheduled Capacities";
  // const workbook = XLSX.readFile(filename);
  // prepareXlsxIterator(workbook, filename, sheetName);
  const stream = fs.createReadStream(filename);
  process_RS(stream, (workbook) => {
    prepareXlsxIterator(workbook, filename, sheetName);
  });
} catch (e) {
  console.log(e);
}
