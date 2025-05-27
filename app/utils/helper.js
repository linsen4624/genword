/* eslint-disable jsdoc/check-tag-names */
/* eslint-disable indent */
//*
// @Autor: Vincent Lin
// @Date: 2025-05-15
// */

const {
  TableCell,
  VerticalAlign,
  Paragraph,
  AlignmentType,
  ImageRun,
  ShadingType,
  TextRun,
  InternalHyperlink,
  Bookmark,
  WidthType,
  TableRow,
  Table,
} = require("docx");
const fs = require("fs");
const { Colors, table_config } = require("./styling");

const for_header = {
  fill: Colors.pink,
  type: ShadingType.CLEAR,
  color: "auto",
};
const for_sub_header = {
  fill: Colors.gray,
  type: ShadingType.CLEAR,
  color: "auto",
};

/**
 * @function getParagraph
 * @param {*} para
 * para = {
 *    txt:     String,
 *    cellType:  String, options are [header, subheader, normal] and normal is the default value
 *    alignment: String, options are [left, center, right] and left is the default value
 *    bookmark:  String, the id of a bookmark
 *    style:     String, an id that defined in styling file
 * }
 ** @return {Paragraph}
 */

function getParagraph(para) {
  let align_way = AlignmentType.LEFT;
  if (para.alignment === "center") align_way = AlignmentType.CENTER;
  if (para.alignment === "right") align_way = AlignmentType.RIGHT;

  if (typeof para.title !== "string") para.title = String(para.title);
  const cType = para.cellType || "normal";

  const txtObj = new TextRun({
    text: para.title || "",
  });

  const bmObj = new Bookmark({
    id: para.bookmark || "",
    children: [txtObj],
  });

  return new Paragraph({
    children: [para.bookmark ? bmObj : txtObj],
    alignment: align_way,
    style:
      para.style ||
      (cType === "header" && "header") ||
      (cType === "subheader" && "sub_header") ||
      null,
  });
}

/**
 * @function getRow
 * @param {*} para
 * para = {
 *    tableHeader:     Object
 *    children:        Array
 * }
 ** @return {TableRow}
 */

function getRow(para) {
  return new TableRow({
    tableHeader: para.tableHeader,
    height: table_config.rowHeight,
    children: para.children,
  });
}

/**
 * @function getCell
 * @param {*} para
 * para = {
 *    borders:   Object
 *    width:     Number
 *    title:     String, required
 *    cellType:  String, options are [header, subheader, normal] and normal is the default value
 *    alignment: String, options are [left, center, right] and left is the default value
 *    bookmark:  String, the id of a bookmark
 *    rows:      Number, the cell's rowSpan
 *    cols:      Number, the cell's colspan
 *    style:     String, an id that defined in styling file
 *    gray_bg    Boolean
 * }
 ** @return {TableCell}
 */

function getCell(para) {
  let shade = null;
  if (para.cellType === "header") shade = for_header;
  if (para.cellType === "subheader" || para.gray_bg) shade = for_sub_header;
  const cellWidth = para.width
    ? { size: para.width, type: WidthType.DXA }
    : { size: 0, type: WidthType.AUTO };

  return new TableCell({
    borders: para.borders,
    width: cellWidth,
    children: [
      getParagraph({
        title: para.title,
        alignment: para.alignment,
        cellType: para.cellType,
        bookmark: para.bookmark,
        style: para.style,
      }),
    ],
    rowSpan: para.rows || null,
    columnSpan: para.cols || null,
    shading: shade,
    verticalAlign: VerticalAlign.CENTER,
  });
}

/**
 * @function getLinkCell
 * @param {*} para
 * para = {
 *    title:     String, required
 *    width:     Number
 *    cellType:  String, options are [header, subheader, normal] and normal is the default value
 *    alignment: String, options are [left, center, right] and left is the default value
 *    target:    String, the id of the element that should be linked to
 *    links:     Object, with title and target keys when more than one link within a cell.
 * }
 ** @return {TableCell}
 */

function getLinkCell(para) {
  let shade = null;
  if (para.cellType === "header") shade = for_header;
  if (para.cellType === "subheader") shade = for_sub_header;

  let align_way = AlignmentType.LEFT;
  if (para.alignment === "center") {
    align_way = AlignmentType.CENTER;
  } else if (para.alignment === "right") {
    align_way = AlignmentType.RIGHT;
  }

  const cellWidth = para.width
    ? { size: para.width, type: WidthType.DXA }
    : { size: 0, type: WidthType.AUTO };

  const single_link = [
    new InternalHyperlink({
      children: [
        new TextRun({
          text: para.title,
          bold: para.cellType !== "normal",
          style: "Hyperlink",
        }),
      ],
      anchor: para.target,
    }),
  ];

  const multiple_links = [];
  if (para.links) {
    para.links.forEach((item) => {
      multiple_links.push(
        new InternalHyperlink({
          children: [
            new TextRun({
              text: item.title,
              bold: para.cellType !== "normal",
              style: "Hyperlink",
            }),
          ],
          anchor: item.target,
        })
      );
      multiple_links.push(new TextRun(" "));
    });
  }
  return new TableCell({
    width: cellWidth,
    children: [
      new Paragraph({
        children: multiple_links.length === 0 ? single_link : multiple_links,
        alignment: align_way,
      }),
    ],
    shading: shade,
    verticalAlign: VerticalAlign.CENTER,
  });
}

/**
 * @function getImageCell
 * @param {*} para
 * para = {
 *    path: String, required
 *    type: String, required
 *    size: Object, with two keys, w and h
 *    cols: Number, the cell's colspan
 * }
 ** @return {TableCell}
 */

function getImageCell(para) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        children: [
          new ImageRun({
            type: para.type,
            data: fs.readFileSync(para.path),
            transformation: {
              width: para.size.w,
              height: para.size.h,
            },
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
    ],
    columnSpan: para.cols || null,
  });
}

/**
 * @function getDynamicTable
 * @param {*} para
 * para = {
 *    category: String
 *    prefix:   String
 *    title:    String
 *    data:     Array
 * }
 ** @return {Table}
 */

function getDynamicTable(para) {
  const DataLists = para.data || [];
  const rows = DataLists.map((item, index) => {
    const numbering = `${para.prefix}.${index + 1}`;
    return new TableRow({
      height: table_config.rowHeight,
      children: [
        new TableCell({
          width: {
            size: 500,
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              children: [
                new Bookmark({
                  id: `${para.category}_${para.prefix}_${index + 1}`,
                  children: [new TextRun({ text: numbering })],
                }),
              ],
            }),
          ],
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          children: [new Paragraph({ text: item })],
          verticalAlign: VerticalAlign.CENTER,
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
      new TableRow({
        height: table_config.rowHeight,
        children: [
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [new TextRun(para.title)],
                alignment: AlignmentType.CENTER,
              }),
            ],
            shading: {
              fill: Colors.gray,
              type: ShadingType.CLEAR,
              color: "auto",
            },
            columnSpan: 2,
          }),
        ],
      }),
      ...rows,
    ],
  });
}

/**
 * @function getPhotosTable
 * @param {*} para
 * para = {
 *    pg:  Array
 * }
 ** @return {Table}
 */

function getPhotosTable(pg) {
  const photos = pg.photos || pg;
  const len = photos?.length || 0;
  if (len <= 0) return;
  const photo_prefix = "images";
  const setPhoto = (p) => {
    return p && p !== ""
      ? new ImageRun({
          type: "jpg",
          data: fs.readFileSync(photo_prefix + p),
          transformation: {
            width: 325,
            height: 250,
            // width: 236,
            // height: 170,
          },
        })
      : new TextRun("NA");
  };

  const photoRows = [];
  if (pg.name && pg.name !== "") {
    photoRows.push(
      new TableRow({
        children: [
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                text: pg.name,
                alignment: AlignmentType.CENTER,
              }),
            ],
            columnSpan: 2,
          }),
        ],
      })
    );
  }
  let image_cells = [];
  let text_cells = [];

  photos.forEach((item, index) => {
    const flag = index + 1;

    image_cells.push(
      new TableCell({
        children: [
          new Paragraph({
            children: [setPhoto(item.url)],
            alignment: AlignmentType.CENTER,
          }),
        ],
      })
    );

    text_cells.push(
      new TableCell({
        children: [
          new Paragraph({
            text: item.description,
            alignment: AlignmentType.CENTER,
          }),
        ],
      })
    );

    if (flag % 2 === 0) {
      photoRows.push(new TableRow({ children: image_cells }));
      photoRows.push(new TableRow({ children: text_cells }));
      image_cells = [];
      text_cells = [];
    } else if (flag === len) {
      image_cells.push(
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              text: "NA",
              alignment: AlignmentType.CENTER,
            }),
          ],
        })
      );
      text_cells.push(
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              text: "-",
              alignment: AlignmentType.CENTER,
            }),
          ],
        })
      );
      photoRows.push(new TableRow({ children: image_cells }));
      photoRows.push(new TableRow({ children: text_cells }));
    }
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: photoRows,
  });
}

/**
 * @function getCleanedString
 * @param {*} str
 * @purpose remove space, punctuation and symbols from a string
 ** @return {String}
 */

function getCleanedString(str) {
  return str.replace(/[^a-zA-Z0-9]/g, "");
}

/**
 * @function getShortString
 * @param {*} str
 * @purpose Truncating the string when it's length longer than a value given.
 ** @return {String}
 */

function getShortString(str, num) {
  if (str.length > num) return str.substring(0, num) + "...";
  return str;
}

/**
 * @function getFormattedTextArray
 * @param {*} str
 * @purpose Marking the content where located in brackets as red and remove the brackets.
 ** @return {Array}
 */

function getFormattedTextArray(str) {
  if (str.indexOf("[") === -1) return [new TextRun(str)];
  const parts = [];
  let currentIndex = 0;

  while (currentIndex < str.length) {
    const openBracketIndex = str.indexOf("[", currentIndex);
    if (openBracketIndex === -1) {
      parts.push(new TextRun(str.substring(currentIndex)));
      break;
    }

    if (openBracketIndex > currentIndex) {
      parts.push(new TextRun(str.substring(currentIndex, openBracketIndex)));
    }

    const closeBracketIndex = str.indexOf("]", openBracketIndex + 1);
    if (closeBracketIndex === -1) {
      parts.push(new TextRun(str.substring(openBracketIndex)));
      break;
    }
    parts.push(
      new TextRun({
        text: str.substring(openBracketIndex + 1, closeBracketIndex),
        color: Colors.red,
      })
    );
    currentIndex = closeBracketIndex + 1;
  }
  return parts;
}

/**
 * @function getFormattedConclusion
 * @param {resultStr, isConclusion}
 ** @return {Array, Object}
 */

function getFormattedConclusion(resultStr, isConclusion) {
  let conclusion_result = "CONFORM";
  let resultColor = Colors.black;
  let conclusion_text = " to client's requirement";
  if (resultStr === "not confirmed") {
    conclusion_result = "NOT CONFORM";
    resultColor = Colors.red;
  }
  if (resultStr === "pending") {
    conclusion_result = "PENDING";
    conclusion_text = " for client's evaluation";
    resultColor = Colors.yellow;
  }

  if (isConclusion) {
    return [
      new TextRun({
        text: conclusion_result,
        bold: true,
        size: 24,
        color: resultColor,
      }),
      new TextRun({
        text: conclusion_text,
        bold: true,
      }),
    ];
  }
  return new TextRun({
    text: conclusion_result,
    bold: true,
    color: resultColor,
  });
}

module.exports = {
  getRow,
  getCell,
  getImageCell,
  getLinkCell,
  getDynamicTable,
  getPhotosTable,
  getCleanedString,
  getShortString,
  getFormattedTextArray,
  getFormattedConclusion,
};
