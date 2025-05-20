const { Table, TableRow, WidthType, Paragraph } = require("docx");
const d = require("../reportData.json");
const { getCell, getPhotosTable } = require("../helper");

const empty_paragraph = new Paragraph("");
const desp =
  "Some abnormal information may not affect the inspection conclusion according to stated requirements but be necessary to report for reference.";

function getNoteLists() {
  return d.OtherNotes.map((item, index) => {
    return new TableRow({
      children: [
        getCell({
          title: `10.${index + 1}`,
          alignment: "center",
        }),
        getCell({
          title: item,
        }),
      ],
    });
  });
}

function getONTable() {
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
            title: "10. Other Note",
            cellType: "subheader",
            alignment: "left",
            cols: 2,
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "Description",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({ title: desp, gray_bg: true }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "No.",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Reference Notes",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),

      ...getNoteLists(),
    ],
  });
}

const ON_Tables = [
  getONTable(),
  empty_paragraph,
  getPhotosTable(["", "", "", ""]),
];

module.exports = ON_Tables;
