const { Table, WidthType, Paragraph } = require("docx");
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
const sn = 2;
const desp =
  "Randomly select samples, proceed tests on site to verify some important characteristics of product, report findings and conformities.";
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
          title: item.name,
          alignment: "center",
        }),
        getCell({
          title: item.Criteria,
          alignment: "center",
        }),
        getCell({
          title: item.SampleSize,
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

function getOSTTable() {
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
            cols: 2,
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
          getCell({ title: desp, cols: 4, gray_bg: true }),
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
            gray_bg: true,
          }),
          getCell({
            title: "Criteria",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Samples",
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

let OST_Tables = [getOSTTable()];

if (photogroup.length > 0) {
  photogroup.forEach((item) => {
    OST_Tables.push(empty_paragraph);
    OST_Tables.push(getPhotosTable(item));
  });
}

if (sap?.length > 0) {
  OST_Tables.push(empty_paragraph);
  OST_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for On Site Test",
      data: sap,
    })
  );
}
if (sap_photos?.length > 0) {
  OST_Tables.push(empty_paragraph);
  OST_Tables.push(getPhotosTable(sap_photos));
}

if (refer?.length > 0) {
  OST_Tables.push(empty_paragraph);
  OST_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for On Site Test",
      data: refer,
    })
  );
}

if (refer_photos?.length > 0) {
  OST_Tables.push(empty_paragraph);
  OST_Tables.push(getPhotosTable(refer_photos));
}

if (dataSheets?.length > 0) {
  OST_Tables = [...OST_Tables, ...getDataSheets(dataSheets)];
}

module.exports = OST_Tables;
