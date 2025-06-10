const { Table, TableRow, WidthType, Paragraph } = require("docx");
const { getCell } = require("../helper");
const Quantity_Table = require("./details/quantity");
const WS_Table = require("./details/workmanship");
const OST_Table = require("./details/onsitetest");
const PDW_Table = require("./details/pdw");
const SMC_Table = require("./details/smc");
const PC_Table = require("./details/productcolor");
const PLM_Table = require("./details/plm");
const SM_Table = require("./details/shippingmark");
const PP_Table = require("./details/pp");

const fs = require("fs");
const { json_target_path } = require("../styling");
const new_json_content = fs.readFileSync(json_target_path, "utf8");
const d = JSON.parse(new_json_content);
if (!d || Object.keys(d).length < 10) return;

const Title_Table = new Table({
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
          title: "INSPECTION DETAILS",
          cellType: "header",
          alignment: "center",
        }),
      ],
    }),
  ],
});

const details = [
  Title_Table,
  new Paragraph(""),
  ...Quantity_Table,
  new Paragraph(""),
  ...WS_Table,
  new Paragraph(""),
  ...OST_Table,
  new Paragraph(""),
  ...PDW_Table,
  new Paragraph(""),
  ...SMC_Table,
  new Paragraph(""),
  ...PC_Table,
  new Paragraph(""),
  ...PLM_Table,
  new Paragraph(""),
  ...SM_Table,
  new Paragraph(""),
  ...PP_Table,
];

module.exports = details;
