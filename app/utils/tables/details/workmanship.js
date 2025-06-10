const {
  Table,
  WidthType,
  Paragraph,
  convertInchesToTwip,
  TableRow,
  TableCell,
  TableBorders,
} = require("docx");

const {
  getRow,
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");
const { table_config, json_target_path } = require("../../styling");
const fs = require("fs");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const empty_paragraph = new Paragraph("");
const sn = 1;
const desp =
  "Randomly draw samples, examine style, construction, color, size, appearance by visual inspection, test the basic function, performance, safety, other characteristics, report defects caused by workmanship.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const checkLists = d.InspectionCategories[sn].checklist;
const photogroup = d.InspectionCategories[sn].PhotoGroup;
const result = d.InspectionCategories[sn].Result;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const sap_photos = d.InspectionCategories[sn].SpecialAttentionPhotos;
const refer_photos = d.InspectionCategories[sn].ReferenceNotePhotos;
const DataLists = d.POItems || [];

function getDefectsTable() {
  const hyphen_cell_width = convertInchesToTwip(0.29);
  let defects = [];
  DataLists.forEach((item) => {
    const defect_title = getRow({
      children: [
        getCell({
          title: `Item No: ${item.ItemNo}, Sample Size= ${item.SampleSize} ${d.ProductUnit}`,
          alignment: "left",
          cols: 5,
        }),
      ],
    });
    const defect_details = item.DefectsList.map((ele) => {
      return getRow({
        children: [
          getCell({ title: "-", width: hyphen_cell_width, alignment: "left" }),
          getCell({ title: ele.defectName, alignment: "left" }),
          getCell({ title: ele.CriticaldefectFounded, alignment: "center" }),
          getCell({ title: ele.MajorDefectFounded, alignment: "center" }),
          getCell({ title: ele.MinorDefectFounded, alignment: "center" }),
        ],
      });
    });

    const dt = d.DefectsTotal || item.DefectsTotal;
    let totals = [];
    if (dt) {
      totals = [
        getRow({
          children: [
            getCell({
              title: "",
              width: hyphen_cell_width,
              alignment: "left",
              gray_bg: true,
            }),
            getCell({
              title: "Total Found",
              alignment: "right",
              gray_bg: true,
            }),
            getCell({
              title: dt.TotalCriticalDefectsFounded,
              alignment: "center",
              style:
                dt.TotalCriticalDefectsFounded > dt.AllowedCriticalDefects
                  ? "red_mark"
                  : null,
            }),
            getCell({
              title: dt.TotalMajorDefectsFounded,
              alignment: "center",
              style:
                dt.TotalMajorDefectsFounded > dt.AllowedMajorDefects
                  ? "red_mark"
                  : null,
            }),
            getCell({
              title: dt.TotalMinorDefectsFounded,
              alignment: "center",
              style:
                dt.TotalMinorDefectsFounded > dt.AllowedMinorDefects
                  ? "red_mark"
                  : null,
            }),
          ],
        }),
        getRow({
          children: [
            getCell({
              title: "",
              width: hyphen_cell_width,
              alignment: "left",
              gray_bg: true,
            }),
            getCell({ title: "Allowed", alignment: "right", gray_bg: true }),
            getCell({
              title: dt.AllowedCriticalDefects,
              alignment: "center",
            }),
            getCell({
              title: dt.AllowedMajorDefects,
              alignment: "center",
            }),
            getCell({
              title: dt.AllowedMinorDefects,
              alignment: "center",
            }),
          ],
        }),
      ];
    }
    if (totals.length > 0) {
      defects = defects.concat([defect_title, ...defect_details, ...totals]);
    } else {
      defects = defects.concat([defect_title, ...defect_details]);
    }
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
            width: convertInchesToTwip(4.92),
            title: "Defects",
            gray_bg: true,
            alignment: "center",
            cols: 2,
          }),
          getCell({
            title: "Critical",
            gray_bg: true,
            alignment: "center",
          }),
          getCell({
            title: "Major",
            gray_bg: true,
            alignment: "center",
          }),
          getCell({
            title: "Minor",
            gray_bg: true,
            alignment: "center",
          }),
        ],
      }),
      ...defects,
    ],
  });
}

function getCheckLists() {
  const check_list_title = getRow({
    children: [
      getCell({
        width: convertInchesToTwip(0.52),
        title: "No.",
        alignment: "left",
        gray_bg: true,
      }),
      getCell({
        width: convertInchesToTwip(1.53),
        title: "Check Point",
        gray_bg: true,
      }),
      getCell({
        width: convertInchesToTwip(4.96),
        title: "Criteria",
        gray_bg: true,
        alignment: "center",
        cols: 4,
      }),
    ],
  });

  const check_list_details = checkLists.map((item, index) => {
    return getRow({
      children: [
        getCell({ title: `${sn + 1}.${index + 1}`, alignment: "left" }),
        getCell({ title: item.name, alignment: "left" }),
        getCell({ title: item.Criteria, alignment: "left", cols: 4 }),
      ],
    });
  });
  return [check_list_title, ...check_list_details];
}

function getDefectPhotos() {
  const defect_photo_table = [];
  const defect_title = new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: [
      getRow({
        children: [
          getCell({
            title: "Defect Photos",
            cellType: "subheader",
            alignment: "center",
          }),
        ],
      }),
    ],
  });

  defect_photo_table.push(defect_title);

  if (photogroup.length > 0) {
    photogroup.forEach((item) => {
      defect_photo_table.push(empty_paragraph);
      defect_photo_table.push(getPhotosTable(item));
    });
  }
  return defect_photo_table;
}

function getPartOne() {
  return new Table({
    borders: { bottom: { style: "none" } },
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
            cols: 4,
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
            width: convertInchesToTwip(0.98),
            title: "Description",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            width: convertInchesToTwip(6.12),
            title: desp,
            gray_bg: true,
            cols: 4,
          }),
        ],
      }),
    ],
  });
}

function getPartTwo() {
  return new Table({
    borders: { bottom: { style: "none" } },
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    margins: table_config.tableMargin,
    rows: [
      getRow({
        children: [
          getCell({
            title: "Sampling Standard",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: d.samplingStandard,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: "Defect",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: "Critical",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: "Major",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: "Minor",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Sampling Plan",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: d.samplingPlan,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: "AQL",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: d["Critical-AQL"],
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: d["Major-AQL"],
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: d["Minor-AQL"],
            cellType: "normal",
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Inspection Level",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: d.inspectionLevel,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: "Sample Size",
            cellType: "normal",
            alignment: "left",
            gray_bg: true,
          }),
          getCell({
            title: d["Critical-SampleSize"],
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: d["Major-SampleSize"],
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: d["Minor-SampleSize"],
            cellType: "normal",
            alignment: "center",
          }),
        ],
      }),
    ],
  });
}

function getPartThree() {
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    margins: table_config.tableMargin,
    rows: getCheckLists(),
  });
}

function getWSTable() {
  return new Table({
    borders: TableBorders.NONE,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: [
      new TableRow({ children: [new TableCell({ children: [getPartOne()] })] }),
      new TableRow({ children: [new TableCell({ children: [getPartTwo()] })] }),
      new TableRow({
        children: [new TableCell({ children: [getPartThree()] })],
      }),
    ],
  });
}

const WS_Tables = [
  getWSTable(),
  empty_paragraph,
  getDefectsTable(),
  empty_paragraph,
  ...getDefectPhotos(),
];

if (sap?.length > 0) {
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Workmanship",
      data: sap,
    })
  );
}

if (sap_photos?.length > 0) {
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(getPhotosTable(sap_photos));
}

if (refer?.length > 0) {
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Workmanship",
      data: refer,
    })
  );
}

if (refer_photos?.length > 0) {
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(getPhotosTable(refer_photos));
}

module.exports = WS_Tables;
