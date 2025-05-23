const { Table, WidthType, Paragraph, BorderStyle } = require("docx");
const d = require("../../reportData.json");
const {
  getRow,
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");
const getDataSheets = require("../datasheets");
const { table_config } = require("../../styling");

const empty_paragraph = new Paragraph("");
const sn = 5;
const desp =
  "Randomly select samples per color to examine / verify the details of color according to the requirement of PO, specifications, client’s comments, claims, approval sample if available, report color issues caused by design, materials, technology.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const result = d.InspectionCategories[sn].Result;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const checkLists = d.InspectionCategories[sn].checklist;
const dataSheets = d.InspectionCategories[sn].datasheet;

function getCheckLists() {
  return checkLists.map((item, index) => {
    return getRow({
      children: [
        getCell({
          title: `${sn + 1}.${index + 1}`,
          alignment: "center",
        }),
        getCell({
          borders: {
            right: {
              style: BorderStyle.NONE,
              size: 0,
              color: "FFFFFF",
            },
          },
          title: item.name,
          alignment: "left",
        }),
        getCell({
          borders: {
            left: {
              style: BorderStyle.NONE,
              size: 0,
              color: "FFFFFF",
            },
          },
          title: item.SampleSize,
          alignment: "left",
        }),
        getCell({
          title: item.Result,
          alignment: "center",
        }),
      ],
    });
  });
}

function getPCTable() {
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
            title: `${sn + 1}. ${subTitle}`,
            cellType: "subheader",
            alignment: "left",
            bookmark: bm,
            cols: 3,
          }),
          getCell({
            title: result,
            alignment: "center",
            style: "red_mark",
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
          getCell({ title: desp, cols: 3, gray_bg: true }),
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
            title: "Check Point",
            cellType: "normal",
            alignment: "center",
            cols: 2,
            gray_bg: true,
          }),
          getCell({
            title: "Result",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),

      ...getCheckLists(),
    ],
  });
}

let PC_Tables = [
  getPCTable(),
  empty_paragraph,
  getPhotosTable(["", "", "", ""]),
];

if (sap?.length > 0) {
  PC_Tables.push(empty_paragraph);
  PC_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Product Color",
      data: sap,
    })
  );
  PC_Tables.push(empty_paragraph);
  PC_Tables.push(getPhotosTable(["", "", "", ""]));
}

if (refer?.length > 0) {
  PC_Tables.push(empty_paragraph);
  PC_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Product Color",
      data: refer,
    })
  );
  PC_Tables.push(empty_paragraph);
  PC_Tables.push(getPhotosTable(["", "", "", ""]));
}

if (dataSheets?.length > 0) {
  PC_Tables = [...PC_Tables, ...getDataSheets(dataSheets)];
}

module.exports = PC_Tables;
