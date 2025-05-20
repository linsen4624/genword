const { Table, TableRow, WidthType, Paragraph } = require("docx");
const d = require("../../reportData.json");
const {
  getCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
} = require("../../helper");

const empty_paragraph = new Paragraph("");
const sn = 0;
const desp =
  "Check / verify the quantity of product available. The standard procedure of pre-shipment inspection required 100% of product has been finished production and ≥80% of product has been packed into export carton.";
const subTitle = d.InspectionCategories[sn].CategoryName;
const result = d.InspectionCategories[sn].Result;
const bm = getCleanedString(subTitle).toLowerCase();
const sap = d.InspectionCategories[sn].SpecialAttention;
const refer = d.InspectionCategories[sn].ReferenceNote;
const DataLists = d.POItems || [];

const Total_Values = {
  POQty: 0,
  ShipmentQtyOfPackage: 0,
  ShipmentQtyOfProduct: 0,
  PackedQtyOfPackage: 0,
  PackedQtyOfProduct: 0,
  SampleCartonCounts: 0,
  SampleSize: 0,
};

function getPOItemRows() {
  return DataLists.map((item) => {
    Total_Values.POQty += item.POQty;
    Total_Values.ShipmentQtyOfPackage += item.ShipmentQtyOfPackage;
    Total_Values.ShipmentQtyOfProduct += item.ShipmentQtyOfProduct;
    Total_Values.PackedQtyOfPackage += item.PackedQtyOfPackage;
    Total_Values.PackedQtyOfProduct += item.PackedQtyOfProduct;
    Total_Values.SampleCartonCounts += item.SampleCartonCounts;
    Total_Values.SampleSize += item.SampleSize;

    return new TableRow({
      children: [
        getCell({
          title: item.PONo,
          alignment: "center",
        }),
        getCell({
          title: item.ItemNo,
          alignment: "center",
        }),
        getCell({
          title: item.POQty,
          alignment: "center",
        }),
        getCell({
          title: item.ShipmentQtyOfPackage,
          alignment: "center",
        }),
        getCell({
          title: item.ShipmentQtyOfProduct,
          alignment: "center",
        }),
        getCell({
          title: item.PackedQtyOfPackage,
          alignment: "center",
        }),
        getCell({
          title: item.PackedQtyOfProduct,
          alignment: "center",
        }),
        getCell({
          title: item.SampleCartonCounts,
          alignment: "center",
        }),
        getCell({
          title: item.SampleSize,
          alignment: "center",
        }),
      ],
    });
  });
}

function getQuantityTable() {
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
            cols: 7,
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
          getCell({ title: desp, cols: 8, gray_bg: true }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: "P.O. No.",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
            rows: 2,
          }),
          getCell({
            title: "Item No.",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
            rows: 2,
          }),
          getCell({
            title: "P.O. Qty",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: "Shipment Qty",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
            cols: 2,
          }),
          getCell({
            title: "Packed Qty",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
            cols: 2,
          }),
          getCell({
            title: "Sample Size",
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
            cols: 2,
          }),
        ],
      }),
      new TableRow({
        children: [
          getCell({
            title: `(${d.ProductUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.ProductUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.PackagingUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.ProductUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.PackagingUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.ProductUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
          getCell({
            title: `(${d.PackagingUnit})`,
            cellType: "normal",
            alignment: "center",
            gray_bg: true,
          }),
        ],
      }),

      ...getPOItemRows(),

      new TableRow({
        children: [
          getCell({
            title: "Total",
            cellType: "normal",
            alignment: "center",
            cols: 2,
            gray_bg: true,
          }),
          getCell({
            title: Total_Values.POQty,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.ShipmentQtyOfPackage,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.ShipmentQtyOfProduct,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.PackedQtyOfPackage,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.PackedQtyOfProduct,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.SampleCartonCounts,
            cellType: "normal",
            alignment: "center",
          }),
          getCell({
            title: Total_Values.SampleSize,
            cellType: "normal",
            alignment: "center",
          }),
        ],
      }),
    ],
  });
}

const photos = getPhotosTable([
  // "images/test/001.jpg",
  // "images/test/000.jpg",
  // "images/test/001.jpg",
  // "images/test/000.jpg",
  "",
  "",
  "",
  "",
]);

const Quantity_Tables = [getQuantityTable(), empty_paragraph, photos];

if (sap?.length > 0) {
  Quantity_Tables.push(empty_paragraph);
  Quantity_Tables.push(
    getDynamicTable({
      category: bm + "_sap",
      prefix: sn + 1,
      title: "Special Attention Point for Quantity",
      data: sap,
    })
  );
  Quantity_Tables.push(empty_paragraph);
  Quantity_Tables.push(getPhotosTable(["", "", "", ""]));
}

if (refer?.length > 0) {
  Quantity_Tables.push(empty_paragraph);
  Quantity_Tables.push(
    getDynamicTable({
      category: bm + "_refer",
      prefix: sn + 1,
      title: "Reference Note for Quantity",
      data: refer,
    })
  );
  Quantity_Tables.push(empty_paragraph);
  Quantity_Tables.push(getPhotosTable(["", "", "", ""]));
}

module.exports = Quantity_Tables;
