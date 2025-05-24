const {
  Table,
  WidthType,
  convertMillimetersToTwip,
  TableRow,
  TableCell,
  ImageRun,
  Paragraph,
} = require("docx");
const fs = require("fs");
const d = require("../reportData.json");
const { getRow, getCell, getShortString } = require("../helper");
const { table_config } = require("../styling");

function getInfoTable() {
  const sub_header_cell_width = convertMillimetersToTwip(40.7);
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
            title: "INSPECTION INFORMATION",
            cellType: "header",
            cols: 4,
            alignment: "center",
          }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Client",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.Client, cols: 3 }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Supplier",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.Supplier, cols: 3 }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Factory",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.Factory, cols: 3 }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "P.O.No.",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: getShortString(d.PoNo, 10) }),
          getCell({
            title: "Quantity",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: `${d.ShipmentQty} ${d.ProductUnit}` }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Item No.",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.ItemNo, cols: 3 }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Product Description",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.ProductDescription, cols: 3 }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Inspection Type",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.InspectionType }),
          getCell({
            title: "Sequence",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.Sequence }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Inspection Date",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.InspectionDate }),
          getCell({ title: "Location", cellType: "subheader" }),
          getCell({ title: d.Location }),
        ],
      }),
      getRow({
        children: [
          getCell({
            title: "Inspection Basis",
            cellType: "subheader",
            width: sub_header_cell_width,
          }),
          getCell({ title: d.InspectionBasis, cols: 3 }),
        ],
      }),
    ],
  });
}

function getPictureTable() {
  const imgs = d.ProductPhotos;
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    margins: table_config.tableMargin,
    rows: [
      new TableRow({
        height: { value: convertMillimetersToTwip(60), rule: "atLeast" },
        children: [
          new TableCell({
            width: {
              size: convertMillimetersToTwip(89),
              type: WidthType.DXA,
            },
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "jpg",
                    data: fs.readFileSync(`images${imgs[0].url}`),
                    transformation: {
                      width: 325,
                      height: 250,
                    },
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: {
              size: convertMillimetersToTwip(89),
              type: WidthType.DXA,
            },
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "jpg",
                    data: fs.readFileSync(`images${imgs[1].url}`),
                    transformation: {
                      width: 325,
                      height: 250,
                    },
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

const Info_Tables = [getInfoTable(), new Paragraph(""), getPictureTable()];

module.exports = Info_Tables;
