const {
  Table,
  TableRow,
  WidthType,
  TableCell,
  VerticalAlign,
  Paragraph,
  convertInchesToTwip,
} = require("docx");
const d = require("../reportData.json");
const {
  getRow,
  getCell,
  getLinkCell,
  getCleanedString,
  getFormattedConclusion,
} = require("../helper");
const { table_config } = require("../styling");
const sub_header_cell_width = convertInchesToTwip(2.75);

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

    return getRow({
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
            title: "INSPECTION RESULT SUMMARY",
            cols: 4,
            bookmark: "summary",
            cellType: "header",
            alignment: "center",
          }),
        ],
      }),
      getRow({
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
        height: { value: convertInchesToTwip(0.48), rule: "atLeast" },
        children: [
          getCell({
            width: sub_header_cell_width,
            title: "OVERALL CONCLUSION",
            cellType: "subheader",
            alignment: "right",
            style: "big_header",
          }),
          new TableCell({
            width: { size: convertInchesToTwip(4.25), type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            columnSpan: 3,
            children: [
              new Paragraph({
                children: getFormattedConclusion(d.Result),
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

module.exports = geSummaryTable;
