const {
  Table,
  TableRow,
  WidthType,
  TableCell,
  VerticalAlign,
  Paragraph,
  convertInchesToTwip,
  AlignmentType,
} = require("docx");
const {
  getRow,
  getCell,
  getLinkCell,
  getCleanedString,
  getFormattedConclusion,
} = require("../helper");
const { table_config, json_target_path } = require("../styling");
const sub_header_cell_width = convertInchesToTwip(2.75);
const all_sap_links = [];
const all_refer_links = [];

const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

function getDataRows() {
  const DataLists = d.InspectionCategories || [];
  return DataLists.map((item, index) => {
    const for_bookmark = getCleanedString(item.CategoryName).toLowerCase();
    const SerialNo = index + 1;
    const sap_links = [];
    const refer_links = [];

    item.SpecialAttention.forEach((ele, idx) => {
      const tmp_sap = {
        title: `${SerialNo}.${idx + 1}`,
        target: for_bookmark + `_sap_${SerialNo}_${idx + 1}`,
      };
      sap_links.push(tmp_sap);
      all_sap_links.push(Object.assign({}, tmp_sap, { text: ele }));
    });

    item.ReferenceNote.forEach((ele, idx) => {
      const tmp_refer = {
        title: `${SerialNo}.${idx + 1}`,
        target: for_bookmark + `_refer_${SerialNo}_${idx + 1}`,
      };
      refer_links.push(tmp_refer);
      all_refer_links.push(Object.assign({}, tmp_refer, { text: ele }));
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
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              style: "red_mark",
              children: [getFormattedConclusion(item.Result, false)],
            }),
          ],
        }),
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

function getSummaryTable() {
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
                children: getFormattedConclusion(d.Result, true),
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function getSAPSummary() {
  const sap_rows = all_sap_links.map((item) => {
    return getRow({
      children: [
        getLinkCell({
          width: 500,
          title: item.title,
          cellType: "normal",
          alignment: "center",
          target: item.target,
        }),
        getCell({
          title: item.text,
        }),
      ],
    });
  });

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
            title: "Special Attention Point Summary",
            cols: 2,
            cellType: "header",
            alignment: "center",
          }),
        ],
      }),
      ...sap_rows,
    ],
  });
}

function getReferSummary() {
  const refer_rows = all_refer_links.map((item) => {
    return getRow({
      children: [
        getLinkCell({
          width: 500,
          title: item.title,
          cellType: "normal",
          alignment: "center",
          target: item.target,
        }),
        getCell({
          title: item.text,
        }),
      ],
    });
  });

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
            title: "Reference Note Summary",
            cols: 2,
            cellType: "header",
            alignment: "center",
          }),
        ],
      }),
      ...refer_rows,
    ],
  });
}

const SUMMARY_Tables = [getSummaryTable()];
if (all_sap_links.length) {
  SUMMARY_Tables.push(new Paragraph(""));
  SUMMARY_Tables.push(getSAPSummary());
}

if (all_refer_links.length) {
  SUMMARY_Tables.push(new Paragraph(""));
  SUMMARY_Tables.push(getReferSummary());
}

module.exports = SUMMARY_Tables;
