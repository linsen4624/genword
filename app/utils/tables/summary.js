const {
  Table,
  TableRow,
  WidthType,
  TableCell,
  VerticalAlign,
  Paragraph,
  TextRun,
} = require("docx");
const d = require("../reportData.json");
const { getCell, getLinkCell, getCleanedString } = require("../helper");
const { Colors } = require("../styling");
const sub_header_cell_width = 4000;

function getDataRows() {
  const DataLists = d.InspectionCategories || [];
  return DataLists.map((item, index) => {
    const for_bookmark = getCleanedString(item.CategoryName).toLowerCase();
    const SerialNo = index + 1;
    const sap_links = [];
    const refer_links = [];

    item.SpecialAttention.forEach((ele, idx) => {
      sap_links.push({
        title: `${SerialNo}.${idx + 1}`,
        target: for_bookmark + `_sap_${SerialNo}_${idx + 1}`,
      });
    });

    item.ReferenceNote.forEach((ele, idx) => {
      refer_links.push({
        title: `${SerialNo}.${idx + 1}`,
        target: for_bookmark + `_refer_${SerialNo}_${idx + 1}`,
      });
    });

    return new TableRow({
      children: [
        getLinkCell({
          width: sub_header_cell_width,
          title: `${SerialNo}. ${item.CategoryName}`,
          cellType: "subheader",
          alignment: "left",
          target: for_bookmark,
        }),
        getCell({ title: item.Result, alignment: "center", style: "red_mark" }),
        getLinkCell({
          cellType: "normal",
          alignment: "center",
          links: sap_links,
          target: for_bookmark,
        }),
        getLinkCell({
          cellType: "normal",
          alignment: "center",
          links: refer_links,
          target: for_bookmark,
        }),
      ],
    });
  });
}

function geSummaryTable() {
  let conclusion_result = "CONFORM";
  let conclusion_text = " to client's requirement";
  if (d.Result === "not confirmed") {
    conclusion_result = "NOT CONFORM";
  }
  if (d.Result === "pending") {
    conclusion_result = "PENDING";
    conclusion_text = " for client's evaluation";
  }
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
            title: "INSPECTION RESULT SUMMARY",
            cols: 4,
            bookmark: "summary",
            cellType: "header",
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "Category",
            cellType: "subheader",
            alignment: "center",
          }),
          getCell({
            title: "Result",
            cellType: "subheader",
            alignment: "center",
          }),
          getCell({
            title: "Special Attention",
            cellType: "subheader",
            alignment: "center",
          }),
          getCell({
            title: "Reference Note",
            cellType: "subheader",
            alignment: "center",
          }),
        ],
      }),

      ...getDataRows(),

      new TableRow({
        height: { value: 700, rule: "exact" },
        children: [
          getCell({
            width: sub_header_cell_width,
            title: "OVERALL CONCLUSION",
            cellType: "subheader",
            alignment: "right",
            style: "big_header",
          }),
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            columnSpan: 3,
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: conclusion_result,
                    bold: true,
                    size: 24,
                    color: Colors.red,
                  }),
                  new TextRun({
                    text: conclusion_text,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

module.exports = geSummaryTable;
