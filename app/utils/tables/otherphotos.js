const { Table, TableRow, WidthType, Paragraph } = require("docx");
// const d = require("../reportData.json");
const { getCell, getPhotosTable } = require("../helper");

const empty_paragraph = new Paragraph("");

function getOPTable() {
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    margins: {
      top: 50,
      bottom: 50,
      left: 100,
      right: 100,
    },
    rows: [
      new TableRow({
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
