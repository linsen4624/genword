const {
  Table,
  TableRow,
  WidthType,
  Paragraph,
  ImageRun,
  TableBorders,
  TableCell,
  VerticalAlign,
} = require("docx");
const fs = require("fs");
const { json_target_path } = require("../styling");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const { getCell } = require("../helper");

const empty_paragraph = new Paragraph("");

function getContractTable() {
  return new Table({
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
          new TableCell({
            rowSpan: 5,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "jpg",
                    data: fs.readFileSync("images/am1.jpg"),
                    transformation: {
                      width: 130,
                      height: 130,
                    },
                  }),
                ],
              }),
            ],
          }),
          getCell({ title: d.AccountManagers[0].name, style: "stress" }),
          new TableCell({
            rowSpan: 5,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "png",
                    data: fs.readFileSync("images/am2.png"),
                    transformation: {
                      width: 130,
                      height: 130,
                    },
                  }),
                ],
              }),
            ],
          }),
          getCell({ title: d.AccountManagers[1].name, style: "stress" }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: d.AccountManagers[0].title }),
          getCell({ title: d.AccountManagers[1].title }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "E: " + d.AccountManagers[0].email }),
          getCell({ title: "E: " + d.AccountManagers[1].email }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Skype: " + d.AccountManagers[0].skype }),
          getCell({ title: "Skype: " + d.AccountManagers[1].skype }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({ text: "M: " + d.AccountManagers[0].mobile }),
            ],
          }),
          new TableCell({
            children: [
              new Paragraph({ text: "M: " + d.AccountManagers[1].mobile }),
            ],
          }),
        ],
      }),
    ],
  });
}

const Back_Cover = new Paragraph({
  children: [
    new ImageRun({
      type: "jpg",
      data: fs.readFileSync("images/logo/cover.jpg"),
      transformation: {
        width: 794,
        height: 1123, // 96 DPI (Standard Screen) Pixels
      },
      floating: {
        horizontalPosition: {
          offset: 0,
        },
        verticalPosition: {
          offset: 0,
        },
        behindDocument: true,
      },
    }),
  ],
});

const Cover_Tables = [Back_Cover];
for (let i = 0; i < 28; i++) {
  Cover_Tables.push(empty_paragraph);
}
Cover_Tables.push(getContractTable());
module.exports = Cover_Tables;
