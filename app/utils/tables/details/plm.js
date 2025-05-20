const { Table, TableRow, WidthType, Paragraph } = require("docx");
const d = require("../../reportData.json");
const {
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");
const getDataSheets = require("../datasheets");

const empty_paragraph = new Paragraph("");
const sn = 6;
const desp =
  "Select samples per item to examine / verify the details of the logo, label, marking, tag, bar code  according to requirements of PO, specifications, client’s comments, claims, approval sample if available, report issues caused by design or wrong pattern used.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const result = d.InspectionCategories[sn].Result;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const checkLists = d.InspectionCategories[sn].checklist;
const dataSheets = d.InspectionCategories[sn].datasheet;

function getCheckLists() {
  return checkLists.map((item, index) => {
    return new TableRow({
      children: [
        getCell({
          title: `${sn + 1}.${index + 1}`,
          alignment: "center",
        }),
        getCell({
          title: item.name,
          alignment: "center",
        }),
        getCell({
          title: item.Result,
          alignment: "center",
        }),
      ],
    });
  });
}

function getPLMTable() {
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
            title: `${sn + 1}. ${subTitle}`,
            cellType: "subheader",
            alignment: "left",
            bookmark: bm,
            cols: 2,
          }),
          getCell({
            title: result,
            alignment: "center",
            style: "red_mark",
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
          getCell({ title: desp, cols: 2, gray_bg: true }),
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
            title: "Check Point",
            cellType: "normal",
            alignment: "center",
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

let PLM_Tables = [
  getPLMTable(),
  empty_paragraph,
  getPhotosTable(["", "", "", ""]),
];

if (sap?.length > 0) {
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Product Label / Marking",
      data: sap,
    })
  );
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(getPhotosTable(["", "", "", ""]));
}

if (refer?.length > 0) {
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Product Label / Marking",
      data: refer,
    })
  );
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(getPhotosTable(["", "", "", ""]));
}

if (dataSheets?.length > 0) {
  PLM_Tables = [...PLM_Tables, ...getDataSheets(dataSheets)];
}

module.exports = PLM_Tables;
