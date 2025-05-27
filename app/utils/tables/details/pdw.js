const { Table, WidthType, Paragraph, convertInchesToTwip } = require("docx");
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
const sn = 3;
const desp =
  "Randomly select and measure and weigh samples, report findings and conformities. In case of no tolerance was specified, adopt general requirement of tolerance of HQTS, or list results for reference.";
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

function getPDWTable() {
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
            width: convertInchesToTwip(5.31),
            title: `${sn + 1}. ${subTitle}`,
            cellType: "subheader",
            alignment: "left",
            bookmark: bm,
            cols: 3,
          }),
          getCell({
            width: convertInchesToTwip(1.69),
            title: result,
            cols: 1,
            alignment: "center",
            style: "red_mark",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            width: convertInchesToTwip(0.88),
            title: "Description",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            width: convertInchesToTwip(6.12),
            title: desp,
            cols: 3,
            gray_bg: true,
          }),
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
            title: "Sample Size",
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

let PDW_Tables = [getPDWTable()];

if (photogroup.length > 0) {
  photogroup.forEach((item) => {
    PDW_Tables.push(empty_paragraph);
    PDW_Tables.push(getPhotosTable(item));
  });
}

if (sap?.length > 0) {
  PDW_Tables.push(empty_paragraph);
  PDW_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Product Dimension & Weight",
      data: sap,
    })
  );
}

if (sap_photos?.length > 0) {
  PDW_Tables.push(empty_paragraph);
  PDW_Tables.push(getPhotosTable(sap_photos));
}

if (refer?.length > 0) {
  PDW_Tables.push(empty_paragraph);
  PDW_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Product Dimension & Weight",
      data: refer,
    })
  );
}

if (refer_photos?.length > 0) {
  PDW_Tables.push(empty_paragraph);
  PDW_Tables.push(getPhotosTable(refer_photos));
}

if (dataSheets?.length > 0) {
  PDW_Tables = [...PDW_Tables, ...getDataSheets(dataSheets)];
}

module.exports = PDW_Tables;
