const {
  PageNumber,
  TextRun,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  ImageRun,
  VerticalAlign,
  TableBorders,
  HeadingLevel,
  BorderStyle,
  ShadingType,
  InternalHyperlink,
  convertInchesToTwip,
} = require("docx");
const fs = require("fs");
const { Colors } = require("../styling");
const { getImage } = require("../helper");
const d = require("../reportData.json");

const signature = new Table({
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  rows: [
    new TableRow({
      height: { value: convertInchesToTwip(0.47), rule: "exact" },
      children: [
        new TableCell({
          width: {
            size: convertInchesToTwip(3.26),
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              children: [new TextRun("Approved by HQTS Supervisor: ")],
              style: "footer_title",
            }),
          ],
          shading: {
            fill: Colors.gray,
            type: ShadingType.CLEAR,
            color: "auto",
          },
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(1.97),
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              children: [
                getImage({
                  type: "jpg",
                  path: d.SupervisorSignPhotoUrl,
                  size: { w: 78, h: 48 },
                  altText: "No Photo Found",
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(0.69),
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun("Date:")],
            }),
          ],
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(1.09),
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun(d.ApprovedDate)],
            }),
          ],
        }),
      ],
    }),
  ],
});

const note = new Paragraph({
  children: [
    new TextRun({
      text: "This report reflects the facts as recorded by HQTS at the time and place of inspection. It does not relieve the manufacturers from their contractual obligations nor prejudice client's right for compensation for any apparent and/or hidden defects not detected during our random inspection or occurring thereafter.",
      italics: true,
    }),
  ],
});

const line = new Paragraph({
  text: "",
  border: {
    top: {
      color: "auto", // Black color (hex code)
      space: 0, // Space between text and line (in points)
      size: 4, // Line thickness (in eighths of a point, e.g., 8 = 1pt)
      style: BorderStyle.SINGLE,
    },
  },
});

const pageinfo = new Table({
  borders: TableBorders.NONE,
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 33.3333,
            type: WidthType.PERCENTAGE,
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [new TextRun("Doc No.: " + d.DocNo)],
              indent: {
                left: 200,
              },
            }),
          ],
        }),
        new TableCell({
          width: {
            size: 33.3333,
            type: WidthType.PERCENTAGE,
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun("V" + d.DocVersion)],
            }),
          ],
        }),
        new TableCell({
          width: {
            size: 33.3333,
            type: WidthType.PERCENTAGE,
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  children: [
                    "Page ",
                    PageNumber.CURRENT,
                    " of ",
                    PageNumber.TOTAL_PAGES,
                  ],
                }),
              ],
              indent: {
                right: 200,
              },
            }),
          ],
        }),
      ],
    }),
  ],
});

const first_page_header = new Table({
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  rows: [
    new TableRow({
      height: { value: convertInchesToTwip(0.79), rule: "atLeast" },
      children: [
        new TableCell({
          width: {
            size: convertInchesToTwip(2.07),
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  type: "png",
                  data: fs.readFileSync("images/logo/logo.png"),
                  transformation: {
                    width: 150,
                    height: 30,
                  },
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(2.66),
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              text: "INSPECTION REPORT",
              heading: HeadingLevel.HEADING_2,
              style: "header_title",
            }),
          ],
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          width: {
            size: convertInchesToTwip(2.28),
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              text: "Report No: " + d.ReportNo,
              alignment: AlignmentType.CENTER,
            }),
          ],
          verticalAlign: VerticalAlign.CENTER,
        }),
      ],
    }),
  ],
});

const header = new Table({
  borders: TableBorders.NONE,
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              text: "Report No: " + d.ReportNo,
              alignment: AlignmentType.RIGHT,
            }),
          ],
        }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new InternalHyperlink({
                  children: [
                    new TextRun({
                      text: "Turn to result summary",
                      bold: true,
                      style: "Hyperlink",
                    }),
                  ],
                  anchor: "summary",
                }),
              ],
              alignment: AlignmentType.RIGHT,
            }),
          ],
        }),
      ],
    }),
  ],
});

const footer = [line, pageinfo];

const first_page_footer = [signature, new Paragraph(""), note, ...footer];

module.exports = {
  footer,
  header,
  first_page_header,
  first_page_footer,
};
