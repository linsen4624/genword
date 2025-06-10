const { Table, TableRow, WidthType, Paragraph } = require("docx");
const { getCell, getPhotosTable } = require("../helper");
const { table_config, json_target_path } = require("../styling");
const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const empty_paragraph = new Paragraph("");

function getOPTable() {
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    margins: table_config.tableMargin,
    rows: [
      new TableRow({
        height: table_config.rowHeight,
        children: [
          getCell({
            title: "11. Other Photos",
            cellType: "subheader",
            alignment: "left",
          }),
        ],
      }),
    ],
  });
}

const OP_Tables = [getOPTable()];
const opg = d.OtherPhotoGroup;
if (opg?.length > 0) {
  opg.forEach((item) => {
    OP_Tables.push(empty_paragraph);
    OP_Tables.push(getPhotosTable(item));
  });
}

module.exports = OP_Tables;
