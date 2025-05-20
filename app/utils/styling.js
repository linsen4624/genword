const { AlignmentType } = require("docx");

const Colors = {
  gray: "f5f5f5",
  pink: "e6a4ae",
  black: "000000",
  red: "8B0000",
};

const styling = {
  default: {
    document: {
      run: {
        size: 20,
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

module.exports = { Colors, styling };
