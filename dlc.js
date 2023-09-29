const xlsx = require("xlsx");
const fse = require("fs-extra");

const workbook = xlsx.readFile("DLC.xlsx");

const sheet_name_list = workbook.SheetNames;
console.log(sheet_name_list);

const processDLCSheet = (sheet) => {
  const jsonWorkbook = xlsx.utils.sheet_to_json(workbook.Sheets[sheet]);

  const listOfFiles = fse.readdirSync("./DLC");
  listOfFiles.forEach((file) => {
    const stringData = fse.readFileSync(`./DLC/${file}`, { encoding: "utf8" });

    const jsonData = JSON.parse(stringData);
    const entries = jsonData.Data.RootChunk.root.Data.entries;

    entries.forEach((entry) => {
      if (
        ((entry.stringId && entry.stringId.length > 0) ||
          (entry.primaryKey && entry.primaryKey.length > 0)) &&
        entry.femaleVariant &&
        entry.femaleVariant.length > 0
      ) {
        // console.log(entry);
        const found = jsonWorkbook.find(
          (r) => r.STRING === entry.stringId || r.STRING === entry.primaryKey
        );
        if (found) {
          if (found["ENG (Female)"] && found["ENG (Female)"].length > 0) {
            entry.femaleVariant = found["ENG (Female)"];
          }
          if (found["ENG (Male)"] && found["ENG (Male)"].length > 0) {
            entry.maleVariant = found["ENG (Male)"];
          }
        }
      }
    });

    fse.outputFileSync(`output/DLC/${file}`, JSON.stringify(jsonData, null, 4));
  });
};

(async () => {
  console.log("Start to transform XLSX to JSON");

  // Only Working with sheet:
  // DLC

  for (let i = 0; i < sheet_name_list.length; i++) {
    const sheet = sheet_name_list[i];
    if (sheet === "DLC") {
      // process sheet onscreens data
      console.log("Start to transform sheet", sheet, "xlsx to json");
      processDLCSheet(sheet);
    }
  }
  console.log("done DLC");
})().catch(console.error);
