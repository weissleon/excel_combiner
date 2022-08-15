import ExcelJS from "exceljs";

async function run() {
  const srcWorkbook = new ExcelJS.Workbook();

  await srcWorkbook.xlsx.readFile("./sample/tale_of_immortal.xlsx");

  const uniqueHeaders: string[] = [];
  for (const worksheet of srcWorkbook.worksheets) {
    console.log("Sheet name:", worksheet.name);
    console.log("Column Count:", worksheet.columnCount);
    console.log("Row Count:", worksheet.rowCount);
    console.log("Dimensions:", worksheet.dimensions);

    const header = worksheet.getRow(1);

    header.eachCell((cell) => {
      const value = cell.value?.toString()!;

      if (!uniqueHeaders.includes(value)) uniqueHeaders.push(value);
    });

    console.log("Headers", uniqueHeaders);

    // const rows = worksheet.getRows(1, worksheet.dimensions.bottom);

    // if (rows === undefined) return;

    // for (const row of rows) {
    //   console.log("Row Number:", row.number);
    //   console.log("Data:", row.values);
    // }
  }
}

run();
