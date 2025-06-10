const {
  Table,
  WidthType,
  Paragraph,
  BorderStyle,
  convertInchesToTwip,
} = require("docx");

const {
  getRow,
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");
const getDataSheets = require("../datasheets");
const { table_config, json_target_path } = require("../../styling");
const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const empty_paragraph = new Paragraph("");
const sn = 7;
const desp =
  "Select export packings to examine / verify the details of shipping mark and labels, markings, other printings at outer packing according to requirements of PO, specifications, client’s comments and other documents report findings and conformities.";
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
          width: convertInchesToTwip(3.29),
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
          width: convertInchesToTwip(0.84),
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

function getSMTable() {
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

let SM_Tables = [getSMTable()];

if (photogroup.length > 0) {
  photogroup.forEach((item) => {
    SM_Tables.push(empty_paragraph);
    SM_Tables.push(getPhotosTable(item));
  });
}

if (sap?.length > 0) {
  SM_Tables.push(empty_paragraph);
  SM_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Shipping Mark",
      data: sap,
    })
  );
}

if (sap_photos?.length > 0) {
  SM_Tables.push(empty_paragraph);
  SM_Tables.push(getPhotosTable(sap_photos));
}

if (refer?.length > 0) {
  SM_Tables.push(empty_paragraph);
  SM_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Shipping Mark",
      data: refer,
    })
  );
}

if (refer_photos?.length > 0) {
  SM_Tables.push(empty_paragraph);
  SM_Tables.push(getPhotosTable(refer_photos));
}

if (dataSheets?.length > 0) {
  SM_Tables = [...SM_Tables, ...getDataSheets(dataSheets)];
}

module.exports = SM_Tables;
