const { Table, WidthType, Paragraph } = require("docx");
const { getRow, getCell, getPhotosTable } = require("../helper");
const { table_config, json_target_path } = require("../styling");
const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const empty_paragraph = new Paragraph("");
const desp =
  "Some abnormal information may not affect the inspection conclusion according to stated requirements but be necessary to report for reference.";

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
