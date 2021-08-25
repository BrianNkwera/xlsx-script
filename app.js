//https://docx.sheetjs.com
const xlsx = require("xlsx");

const workBook = xlsx.readFile("file2.xlsx");

//console.log(workBook.SheetNames);-logs the tabs(sheets) in the file

const sheet = workBook.Sheets["Sheet1"];

const data = xlsx.utils.sheet_to_json(sheet);

// for(let i = 0; i < 330; i ++){

// console.log(data[i].Text.includes("ujauzito"));
// }
data[330].Text.toString();

for (let i = 0; i < data.length; i++) {
  const element = data[i];
  if (typeof element.Text !== "string") {
    continue;
  }

  if (
    element.Text.includes("ujauzito") ||
    element.Text.includes("mjamzito") ||
    element.Text.includes("fungua") ||
    element.Text.includes("fungulia") ||
    element.Text.includes("kliniki")
  ) {
    element.Topic_id = "A1";
  } else if (
    element.Text.includes("tembezi") ||
    element.Text.includes("pumzika")
  ) {
    element.Topic_id = "A1,A3,A4,A5,A7";
  } else if (
    element.Text.includes("dawa") ||
    element.Text.includes("uboreshwaji") ||
    element.Text.includes("namba")
  ) {
    element.Topic_id = "A1,A2,A3,A4,A5,A6,A7";
  } else {
    element.Topic_id = "";
  }
}

console.log(data);

const newWorkBook = xlsx.utils.book_new();
const newSheet = xlsx.utils.json_to_sheet(data);
xlsx.utils.book_append_sheet(newWorkBook,newSheet,"new WorkSheet");

xlsx.writeFile(newWorkBook, "file2_new.xlsx");

