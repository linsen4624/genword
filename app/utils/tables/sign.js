const {
  Table,
  WidthType,
  TableCell,
  Paragraph,
  TextRun,
  AlignmentType,
  convertInchesToTwip,
  VerticalAlign,
} = require("docx");
const { getRow, getCell } = require("../helper");
const { table_config, json_target_path } = require("../styling");
const fixed_width = convertInchesToTwip(1.75);

const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const Sign_Table = new Table({
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  margins: table_config.tableMargin,
  rows: [
    new getRow({
      children: [
        getCell({
          width: fixed_width,
          title: "Inspector:",
          cellType: "subheader",
          alignment: "left",
        }),
        getCell({
          width: fixed_width,
          title: d.Inspector,
        }),
        getCell({
          width: fixed_width,
          title: "Auditor:",
          cellType: "subheader",
          alignment: "center",
        }),
        getCell({
          width: fixed_width,
          title: d.Auditor,
        }),
      ],
    }),
  ],
});

const End_Table = new Table({
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  margins: table_config.tableMargin,
  rows: [
    new getRow({
      children: [
        getCell({
          width: convertInchesToTwip(4.33),
          title: "",
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(2.36),
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "End of Report",
                  bold: true,
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
        getCell({
          width: convertInchesToTwip(4.33),
          title: "",
        }),
      ],
    }),
  ],
});

const SIGN_Tables = [
  Sign_Table,
  new Paragraph({ text: "", spacing: { line: 0 } }),
  End_Table,
];

module.exports = SIGN_Tables;
