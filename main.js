const sanitizeHtml = require("sanitize-html");
class DigitalLib {
  constructor() {
    this.xlsx = require("xlsx");
    this.excel = require("excel4node");
    this.path = require("path");
    this.fs = require("fs");
    this.filesRoutes = [];
  }

  getLetterSlice(c1 = "a", c2 = "z") {
    c1 = c1.toLowerCase();
    c2 = c2.toLowerCase();
    const a = "abcdefghijklmnopqrstuvwxyz".split("");
    return a.slice(a.indexOf(c1), a.indexOf(c2) + 1);
  }

  getAllFileRoutes(directory) {
    this.fs.readdirSync(directory).forEach((file) => {
      const absolute = this.path.join(directory, file);
      if (this.fs.statSync(absolute).isDirectory())
        return this.getAllFileRoutes(absolute);
      else return this.filesRoutes.push(absolute);
    });
  }
  getJsonFromExcelAll(fileRoute, opts = { header: "A" }) {
    let workbook = this.xlsx.readFile(fileRoute);
    let sheet_name_list = workbook.SheetNames;
    let json = [];
    sheet_name_list.forEach((sheet_name, index) => {
      json.push({
        name: sheet_name,
        index: index,
        json: this.xlsx.utils.sheet_to_json(
          workbook.Sheets[sheet_name_list[0]],
          opts
        ),
      });
    });
  }
  getJsonFromExcelFirst(fileRoute, opts = { header: "A" }) {
    let workbook = this.xlsx.readFile(fileRoute);
    let sheet_name_list = workbook.SheetNames;
    return this.xlsx.utils.sheet_to_json(
      workbook.Sheets[sheet_name_list[0]],
      opts
    );
  }
}

const DL = new DigitalLib();

DL.getAllFileRoutes("./maps/");

const json = DL.getJsonFromExcelFirst(DL.filesRoutes[0]);

console.log(DL.getJsonFromExcelFirst(DL.filesRoutes[0]));
