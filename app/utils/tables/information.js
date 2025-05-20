const { Table, TableRow, WidthType } = require("docx");
const d = require("../reportData.json");
const { getCell, getImageCell, getShortString } = require("../helper");

function getInfoTable() {
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
            title: "INSPECTION INFORMATION",
            cellType: "header",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Client", cellType: "subheader" }),
          getCell({ title: d.Client, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Supplier", cellType: "subheader" }),
          getCell({ title: d.Supplier, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Factory", cellType: "subheader" }),
          getCell({ title: d.Factory, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "P.O.No.", cellType: "subheader" }),
          getCell({ title: getShortString(d.PoNo, 10) }),
          getCell({ title: "Quantity", cellType: "subheader" }),
          getCell({ title: `${d.ShipmentQty} ${d.ProductUnit}` }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Item No.", cellType: "subheader" }),
          getCell({ title: d.ItemNo, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Product Description", cellType: "subheader" }),
          getCell({ title: d.ProductDescription, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Inspection Type", cellType: "subheader" }),
          getCell({ title: d.InspectionType }),
          getCell({ title: "Sequence", cellType: "subheader" }),
          getCell({ title: d.Sequence }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Inspection Date", cellType: "subheader" }),
          getCell({ title: d.InspectionDate }),
          getCell({ title: "Location", cellType: "subheader" }),
          getCell({ title: d.Location }),
        ],
      }),
      new TableRow({
        children: [
          getCell({ title: "Inspection Basis", cellType: "subheader" }),
          getCell({ title: d.InspectionBasis, cols: 3 }),
        ],
      }),
      new TableRow({
        children: [
          getImageCell({
            type: "jpg",
            path: "images/test/001.jpg",
            size: { w: 325, h: 250 },
            cols: 2,
          }),
          getImageCell({
            type: "jpg",
            path: "images/test/000.jpg",
            size: { w: 325, h: 250 },
            cols: 2,
          }),
        ],
      }),
    ],
  });
}

module.exports = getInfoTable;
