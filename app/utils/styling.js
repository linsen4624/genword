const { AlignmentType, convertMillimetersToTwip } = require("docx");

const upzip_target_path = "images/ReportPhoto";
const json_target_path = "app/utils/reportData.json";

const Colors = {
  gray: "f2f2f2",
  pink: "e5b7b7",
  black: "000000",
  red: "bd0002",
  yellow: "ffff00",
};

const table_config = {
  tableMargin: {
    top: 0,
    bottom: 0,
    left: 100,
    right: 100,
  },
  rowHeight: { value: convertMillimetersToTwip(5), rule: "atLeast" },
};

const styling = {
  default: {
    document: {
      run: {
        size: 20, // 微软特殊的数字单位Twip，20 相当于10号字
        font: "Arial",
      },
    },
    hyperlink: {
      run: {
        color: Colors.black,
        underline: {
          color: Colors.black,
        },
      },
    },
  },
  paragraphStyles: [
    {
      id: "header_title",
      name: "Header Title",
      run: {
        bold: true,
        size: 32,
        color: Colors.black,
      },
      paragraph: {
        alignment: AlignmentType.CENTER,
      },
    },
    {
      id: "footer_title",
      name: "Footer Title",
      run: {
        bold: true,
        size: 24,
        color: Colors.black,
      },
      paragraph: {
        alignment: AlignmentType.CENTER,
      },
    },
    {
      id: "header",
      name: "Header",
      run: {
        bold: true,
        size: 20,
        color: Colors.black,
      },
    },
    {
      id: "sub_header",
      name: "Sub Header",
      run: {
        bold: true,
        size: 20,
        color: Colors.black,
      },
    },
    {
      id: "big_header",
      name: "Big Header",
      run: {
        bold: true,
        size: 24,
        color: Colors.black,
      },
    },
    {
      id: "red_mark",
      name: "Red Mark",
      run: {
        bold: true,
        color: Colors.red,
      },
    },
    {
      id: "stress",
      name: "Stress",
      run: {
        bold: true,
      },
    },
  ],
};

module.exports = {
  Colors,
  styling,
  table_config,
  upzip_target_path,
  json_target_path,
};
