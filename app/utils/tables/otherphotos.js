const { Table, TableRow, WidthType, Paragraph } = require("docx");
const d = require("../reportData.json");
const { getCell, getPhotosTable } = require("../helper");
const { table_config } = require("../styling");

const empty_paragraph = new Paragraph("");

if (!d || Object.keys(d).length < 10) return;

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
