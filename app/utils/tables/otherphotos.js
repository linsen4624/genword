const { Table, TableRow, WidthType, Paragraph } = require("docx");
// const d = require("../reportData.json");
const { getCell, getPhotosTable } = require("../helper");
const { table_config } = require("../styling");

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

const OP_Tables = [
  getOPTable(),
  empty_paragraph,
  getPhotosTable(["", "", "", ""]),
];

module.exports = OP_Tables;
