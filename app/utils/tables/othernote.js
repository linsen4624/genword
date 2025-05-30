const { Table, WidthType, Paragraph } = require("docx");
const d = require("../reportData.json");
const { getRow, getCell, getPhotosTable } = require("../helper");
const { table_config } = require("../styling");

const empty_paragraph = new Paragraph("");
const desp =
  "Some abnormal information may not affect the inspection conclusion according to stated requirements but be necessary to report for reference.";

if (!d || Object.keys(d).length < 10) return;

function getNoteLists() {
  return d.OtherNotes.map((item, index) => {
    return getRow({
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
    margins: table_config.tableMargin,
    rows: [
      getRow({
        children: [
          getCell({
            title: "10. Other Note",
            cellType: "subheader",
            alignment: "left",
            cols: 2,
          }),
        ],
      }),
      getRow({
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
      getRow({
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
  getPhotosTable(d.OtherNotesPhotos),
];

module.exports = ON_Tables;
