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
const sn = 6;
const desp =
  "Select samples per item to examine / verify the details of the logo, label, marking, tag, bar code  according to requirements of PO, specifications, client’s comments, claims, approval sample if available, report issues caused by design or wrong pattern used.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const result = d.InspectionCategories[sn].Result;
const photogroup = d.InspectionCategories[sn].PhotoGroup;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const sap_photos = d.InspectionCategories[sn].SpecialAttentionPhotos;
const refer_photos = d.InspectionCategories[sn].ReferenceNotePhotos;
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

function getPLMTable() {
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

let PLM_Tables = [getPLMTable()];

if (photogroup.length > 0) {
  photogroup.forEach((item) => {
    PLM_Tables.push(empty_paragraph);
    PLM_Tables.push(getPhotosTable(item));
  });
}

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
}

if (sap_photos?.length > 0) {
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(getPhotosTable(sap_photos));
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
}

if (refer_photos?.length > 0) {
  PLM_Tables.push(empty_paragraph);
  PLM_Tables.push(getPhotosTable(refer_photos));
}

if (dataSheets?.length > 0) {
  PLM_Tables = [...PLM_Tables, ...getDataSheets(dataSheets)];
}

module.exports = PLM_Tables;
