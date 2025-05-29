const {
  Table,
  WidthType,
  Paragraph,
  TableCell,
  VerticalAlign,
  AlignmentType,
} = require("docx");
const {
  getRow,
  getCell,
  getImageCell,
  getFormattedTextArray,
} = require("../helper");
const { table_config } = require("../styling");

function getNormalDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return getRow({
      children: [
        getCell({
          title: item.ItemNo,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({
          title: item.Specification,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({
          title: item.Tolerance,
          cellType: "normal",
          alignment: "center",
        }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: getFormattedTextArray(item.Result),
            }),
          ],
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "normal",
            cols: 4,
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Item No.",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Specification",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Tolerance",
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
      ...rows,
    ],
  });
}

function getWithPicDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return getRow({
      children: [
        getCell({
          title: item.Checkpoint,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({
          title: item.Specification,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({
          title: item.Tolerance,
          cellType: "normal",
          alignment: "center",
        }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: getFormattedTextArray(item.Result),
            }),
          ],
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
            title: data.name,
            cellType: "normal",
            cols: 4,
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Item No: " + data.ItemNo,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getImageCell({
            type: "png",
            path: data.photo.url,
            size: { w: 325, h: 250 },
            cols: 4,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Check Point",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Specification",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Tolerance",
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
      ...rows,
    ],
  });
}

function getCDFDatasheet(data) {
  const rows = data?.datas.map((item, index) => {
    return getRow({
      children: [
        getCell({ title: index + 1, cellType: "normal", alignment: "center" }),
        getCell({
          title: item.ComponentName,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({ title: item.OnCDF, cellType: "normal", alignment: "center" }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: getFormattedTextArray(item.Findings),
            }),
          ],
        }),
        getCell({ title: item.Result, cellType: "normal" }),
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "normal",
            cols: 5,
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: `Item No.: ${data.ItemNo}`,
            cellType: "normal",
            cols: 2,
          }),
          getCell({
            title: `Manufacture Model No.: ${data.Model}`,
            cellType: "normal",
          }),
          getCell({
            title: `Report No.: ${data.ReportNo}`,
            cellType: "normal",
            cols: 2,
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
            title: "Component Name",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "On CDF",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Findings",
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
      ...rows,
    ],
  });
}

function getBarCodeDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return getRow({
      children: [
        getCell({
          title: item.Position,
          cellType: "normal",
          alignment: "center",
        }),
        getCell({
          title: item.Specification,
          cellType: "normal",
          alignment: "center",
        }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: getFormattedTextArray(item.Findings),
            }),
          ],
        }),
        getCell({
          title: item.Result,
          cellType: "normal",
          alignment: "center",
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "normal",
            cols: 4,
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: `Item No.: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: `Color: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: `Size: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Position",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Specification",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Findings",
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
      ...rows,
    ],
  });
}

function getShoesDatasheet(data) {
  const dts = data?.datas;
  const rows = [];
  dts.forEach((item) => {
    if (item.ItemNo) {
      rows.push(
        getRow({
          children: [
            getCell({
              title: "Item No: " + item.ItemNo,
              cellType: "normal",
              alignment: "center",
            }),
          ],
        })
      );
    }

    item?.photos.forEach((pic) => {
      rows.push(
        getRow({
          children: [
            getImageCell({
              type: "jpg",
              path: pic.url,
              // size: { w: 325, h: 250 },
              size: { w: 470, h: 340 },
            }),
          ],
        })
      );
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
            title: data.name,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),
      ...rows,
    ],
  });
}

function getDataSheets(ds) {
  const dss = [];
  ds.forEach((item) => {
    dss.push(new Paragraph(""));
    switch (item.type) {
      case "ShoesDataSheet":
        dss.push(getShoesDatasheet(item));
        break;
      case "WithPicDataSheet":
        dss.push(getWithPicDatasheet(item));
        break;
      case "CDFDataSheet":
        dss.push(getCDFDatasheet(item));
        break;
      case "BarCodeDataSheet":
        dss.push(getBarCodeDatasheet(item));
        break;
      default:
        dss.push(getNormalDatasheet(item));
        break;
    }
  });
  return dss;
}

module.exports = getDataSheets;
