const {
  Table,
  TableRow,
  WidthType,
  TableCell,
  Paragraph,
  TextRun,
  AlignmentType,
  TableBorders,
} = require("docx");
const d = require("../reportData.json");
const { getCell } = require("../helper");

const Sign_Table = new Table({
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
          title: "Inspector:",
          cellType: "subheader",
          alignment: "left",
        }),
        getCell({
          title: d.Inspector,
        }),
        getCell({
          title: "Auditor:",
          cellType: "subheader",
          alignment: "center",
        }),
        getCell({
          title: d.Auditor,
        }),
      ],
    }),
  ],
});

const End_Table = new Table({
  borders: TableBorders.NONE,
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
          title: "",
        }),
        new TableCell({
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
          title: "",
        }),
      ],
    }),
  ],
});

const SIGN_Tables = [Sign_Table, new Paragraph(""), End_Table];

module.exports = SIGN_Tables;
