const { Table, TableRow, WidthType, Paragraph } = require("docx");
const d = require("../../reportData.json");
const {
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");

const empty_paragraph = new Paragraph("");
const sn = 1;
const desp =
  "Randomly draw samples, examine style, construction, color, size, appearance by visual inspection, test the basic function, performance, safety, other characteristics, report defects caused by workmanship.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const checkLists = d.InspectionCategories[sn].checklist;
const result = d.InspectionCategories[sn].Result;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const DataLists = d.POItems || [];

function getDefectsTable() {
  let defects = [];
  DataLists.forEach((item) => {
    const defect_title = new TableRow({
      children: [
        getCell({
          title: `Item No: ${item.ItemNo}, Sample Size= ${item.SampleSize} ${d.ProductUnit}`,
          alignment: "left",
          cols: 5,
        }),
      ],
    });
    const defect_details = item.DefectsList.map((ele) => {
      return new TableRow({
        children: [
          getCell({ title: "-", width: 150, alignment: "left" }),
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
        new TableRow({
          children: [
            getCell({
              title: "",
              width: 150,
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
        new TableRow({
          children: [
            getCell({
              title: "",
              width: 150,
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
            width: 7050,
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
  const check_list_title = new TableRow({
    children: [
      getCell({
        title: "No.",
        alignment: "center",
        gray_bg: true,
      }),
      getCell({
        title: "Check Point",
        gray_bg: true,
      }),
      getCell({
        title: "Criteria",
        gray_bg: true,
        alignment: "center",
        cols: 4,
      }),
    ],
  });

  const check_list_details = checkLists.map((item, index) => {
    return new TableRow({
      children: [
        getCell({ title: `${sn + 1}.${index + 1}`, alignment: "center" }),
        getCell({ title: item.name, alignment: "center" }),
        getCell({ title: item.Criteria, alignment: "center", cols: 4 }),
      ],
    });
  });
  return [check_list_title, ...check_list_details];
}

function getDefectPhotos() {
  const defect_photo_title = new Table({
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
            title: "Defect Photos",
            cellType: "subheader",
            alignment: "center",
          }),
        ],
      }),
    ],
  });

  const defect_photo_list = getPhotosTable([
    // "images/test/001.jpg",
    // "images/test/000.jpg",
    // "images/test/001.jpg",
    // "images/test/000.jpg",
    "",
    "",
    "",
    "",
  ]);
  return [defect_photo_title, empty_paragraph, defect_photo_list];
}

function getWSTable() {
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
            cols: 4,
          }),
          getCell({
            title: result,
            cols: 2,
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
          getCell({ title: desp, cols: 5, gray_bg: true }),
        ],
      }),
      new TableRow({
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
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Critical",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Major",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Minor",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "Sampling Plan",
            cellType: "normal",
            alignment: "center",
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
            alignment: "center",
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

      new TableRow({
        children: [
          getCell({
            title: "Inspection Level",
            cellType: "normal",
            alignment: "center",
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
            alignment: "center",
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
      ...getCheckLists(),
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
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(getPhotosTable(["", "", "", ""]));
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
  WS_Tables.push(empty_paragraph);
  WS_Tables.push(getPhotosTable(["", "", "", ""]));
}

module.exports = WS_Tables;
