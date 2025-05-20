const { Table, TableRow, WidthType, Paragraph } = require("docx");
const { getCell, getImageCell, getPhotosTable } = require("../helper");

function getNormalDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return new TableRow({
      children: [
        getCell({ title: item.ItemNo, cellType: "normal" }),
        getCell({ title: item.Specification, cellType: "normal" }),
        getCell({ title: item.Tolerance, cellType: "normal" }),
        getCell({ title: item.Result, cellType: "normal" }),
      ],
    });
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "subheader",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Item No.", cellType: "subheader" }),
          getCell({ title: "Specification", cellType: "subheader" }),
          getCell({ title: "Tolerance", cellType: "subheader" }),
          getCell({ title: "Result", cellType: "subheader" }),
        ],
      }),
      ...rows,
    ],
  });
}

function getWithPicDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return new TableRow({
      children: [
        getCell({ title: item.Checkpoint, cellType: "normal" }),
        getCell({ title: item.Specification, cellType: "normal" }),
        getCell({ title: item.Tolerance, cellType: "normal" }),
        getCell({ title: item.Result, cellType: "normal" }),
      ],
    });
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
            title: data.name,
            cellType: "subheader",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "Item No: " + data.ItemNo,
            cellType: "subheader",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getImageCell({
            type: "jpg",
            path: "images/test/000.jpg",
            size: { w: 325, h: 250 },
            cols: 4,
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Check Point", cellType: "subheader" }),
          getCell({ title: "Specification", cellType: "subheader" }),
          getCell({ title: "Tolerance", cellType: "subheader" }),
          getCell({ title: "Result", cellType: "subheader" }),
        ],
      }),
      ...rows,
    ],
  });
}

function getCDFDatasheet(data) {
  const rows = data?.datas.map((item, index) => {
    return new TableRow({
      children: [
        getCell({ title: index + 1, cellType: "normal" }),
        getCell({ title: item.ComponentName, cellType: "normal" }),
        getCell({ title: item.OnCDF, cellType: "normal" }),
        getCell({ title: item.Findings, cellType: "normal" }),
        getCell({ title: item.Result, cellType: "normal" }),
      ],
    });
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "subheader",
            cols: 5,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
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
      new TableRow({
        children: [
          getCell({ title: "No.", cellType: "subheader" }),
          getCell({ title: "Component Name", cellType: "subheader" }),
          getCell({ title: "On CDF", cellType: "subheader" }),
          getCell({ title: "Findings", cellType: "subheader" }),
          getCell({ title: "Result", cellType: "subheader" }),
        ],
      }),
      ...rows,
    ],
  });
}

function getBarCodeDatasheet(data) {
  const rows = data?.datas.map((item) => {
    return new TableRow({
      children: [
        getCell({ title: item.Position, cellType: "normal" }),
        getCell({ title: item.Specification, cellType: "normal" }),
        getCell({ title: item.Findings, cellType: "normal" }),
        getCell({ title: item.Result, cellType: "normal" }),
      ],
    });
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
        tableHeader: true,
        children: [
          getCell({
            title: data.name,
            cellType: "subheader",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: `Item No.: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: `Color: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: `Size: ${data.ItemNo}`,
            cellType: "normal",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Position", cellType: "subheader" }),
          getCell({ title: "Specification", cellType: "subheader" }),
          getCell({ title: "Findings", cellType: "subheader" }),
          getCell({ title: "Result", cellType: "subheader" }),
        ],
      }),
      ...rows,
    ],
  });
}

function getShoesDatasheet() {
  return getPhotosTable(["", "", "", ""]);
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
