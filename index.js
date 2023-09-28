const xlsx = require("xlsx");
const fse = require("fs-extra");

const workbook = xlsx.readFile("Update_Only_Full.xlsx");

const sheet_name_list = workbook.SheetNames;
console.log(sheet_name_list);

const processOnscreensSheet = (sheet) => {
  const jsonWorkbook = xlsx.utils.sheet_to_json(workbook.Sheets[sheet]);
  const stringData = fse.readFileSync(
    "./input/en-us/onscreens/onscreens.json.json",
    { encoding: "utf8" }
  );
  const jsonData = JSON.parse(stringData);
  const entries = jsonData.Data.RootChunk.root.Data.entries;

  jsonWorkbook.forEach((entry) => {
    const stringId = entry.STRING;
    const indexOfEntryInJson = entries.findIndex(
      (e) => e.primaryKey === stringId
    );
    try {
      if (indexOfEntryInJson >= 0) {
        if (
          entry["ENG (Female) Merge"] &&
          entry["ENG (Female) Merge"].length > 0
        ) {
          // translated => replace VIE to the femaleVariant
          entries[indexOfEntryInJson].femaleVariant =
            entry["ENG (Female) Merge"];
        }
        if (entry["ENG (Male)"] && entry["ENG (Male)"].length > 0) {
          // translated => replace VIE to the maleVariant
          entries[indexOfEntryInJson].maleVariant = entry["ENG (Male)"];
        }
      }
    } catch (err) {
      console.log(err);
      console.log(entry);
    }
  });

  // replace old entries with new entries
  jsonData.Data.RootChunk.root.Data.entries = entries;

  fse.outputFileSync(
    "output/en-us/onscreens/onscreens.json.json",
    JSON.stringify(jsonData, null, 4)
  );
  console.log("done", sheet);
};

const processSubtitlesSheet = (sheet) => {
  const jsonWorkbook = xlsx.utils.sheet_to_json(workbook.Sheets[sheet]);
  const listOfFolder = fse.readdirSync("./input/en-us/subtitles/"+sheet);

  listOfFolder.forEach((folder) => { //animated
    if (!folder.includes(".json")) {
      const files = fse.readdirSync("./input/en-us/subtitles/"+sheet+'/'+folder);
      files.forEach(file => { //json
        const stringData = fse.readFileSync(
          `./input/en-us/subtitles/${sheet}/${folder}/${file}`,
          { encoding: "utf8" }
        );
        const jsonData = JSON.parse(stringData);
        // entries of official file
        const entries = jsonData.Data.RootChunk.root.Data.entries;
        jsonWorkbook.forEach((entry, index) => {
          if(index === 0) {
            // console.log(jsonWorkbook[0])
          }
          const stringId = entry.STRING;
          const indexOfEntryInJson = entries.findIndex(
            (e) => e.stringId === stringId
          );
          try {
            if (indexOfEntryInJson >= 0) {
              if (
                entry["ENG (Female)"] &&
                entry["ENG (Female)"].length > 0
              ) {
                // translated => replace VIE to the femaleVariant
                entries[indexOfEntryInJson].femaleVariant =
                  entry["ENG (Female)"];
              }
              if (entry["ENG (Male)"] && entry["ENG (Male)"].length > 0) {
                // translated => replace VIE to the maleVariant
                entries[indexOfEntryInJson].maleVariant = entry["ENG (Male)"];
              }
            }
          } catch (err) {
            console.log(err);
            console.log(entry);
          }
        });

        // replace old entries with new entries
        jsonData.Data.RootChunk.root.Data.entries = entries;

        let outputFileName = file;
        if (outputFileName.includes('.json.json')) {
          // outputFileName = outputFileName.replace('.json.json','.json')
        }
        fse.outputFileSync(
          `./output/en-us/subtitles/${sheet}/${folder}/${outputFileName}`,
          JSON.stringify(jsonData, null, 4)
        );
        console.log("done", `${sheet}/${folder}/${outputFileName}`);
      });
    } else {
      
      let singleFileName = folder
      
      const stringData = fse.readFileSync(
        `./input/en-us/subtitles/${sheet}/${singleFileName}`,
        { encoding: "utf8" }
      );
      const jsonData = JSON.parse(stringData);
      // entries of official file
      const entries = jsonData.Data.RootChunk.root.Data.entries;
      jsonWorkbook.forEach((entry, index) => {
        if(index === 0) {
          // console.log(jsonWorkbook[0])
        }
        const stringId = entry.STRING;
        const indexOfEntryInJson = entries.findIndex(
          (e) => e.stringId === stringId
        );
        try {
          if (indexOfEntryInJson >= 0) {
            if (
              entry["ENG (Female)"] &&
              entry["ENG (Female)"].length > 0
            ) {
              // translated => replace VIE to the femaleVariant
              entries[indexOfEntryInJson].femaleVariant =
                entry["ENG (Female)"];
            }
            if (entry["ENG (Male)"] && entry["ENG (Male)"].length > 0) {
              // translated => replace VIE to the maleVariant
              entries[indexOfEntryInJson].maleVariant = entry["ENG (Male)"];
            }
          }
        } catch (err) {
          console.log(err);
          console.log(entry);
        }
      });

      // replace old entries with new entries
      jsonData.Data.RootChunk.root.Data.entries = entries;

      let outputFileName = singleFileName;
      if (outputFileName.includes('.json.json')) {
        // outputFileName = outputFileName.replace('.json.json','.json')
      }
      fse.outputFileSync(
        `./output/en-us/subtitles/${sheet}/${outputFileName}`,
        JSON.stringify(jsonData, null, 4)
      );
      console.log("done", `${sheet}/${outputFileName}`);
    }

  });

  // const stringData = fse.readFileSync(
  //   "./input/en-us/subtitles/media",
  //   { encoding: "utf8" }
  // );
  // const jsonData = JSON.parse(stringData);
  // const entries = jsonData.Data.RootChunk.root.Data.entries;
};

(async () => {
  console.log("Start to transform XLSX to JSON");

  // Only Working with sheet:
  // onscreens
  // media
  // open_world
  // quest

  for (let i = 0; i < sheet_name_list.length; i++) {
    const sheet = sheet_name_list[i];
    if (sheet === "onscreens") {
      // process sheet onscreens data
      console.log("Start to transform sheet", sheet, "xlsx to json");
      processOnscreensSheet(sheet);
    }
    if (sheet === "open_world") {
      // process sheet onscreens data
      console.log("Start to transform sheet", sheet, "xlsx to json");
      processSubtitlesSheet(sheet);
    }
    if (sheet === "media") {
      // process sheet onscreens data
      console.log("Start to transform sheet", sheet, "xlsx to json");
      processSubtitlesSheet(sheet);
    }
    if (sheet === "quest") {
      // process sheet onscreens data
      console.log("Start to transform sheet", sheet, "xlsx to json");
      processSubtitlesSheet(sheet);

    }
  }
})().catch(console.error);
