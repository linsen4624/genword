"use strict";
const Service = require("egg").Service;
const fs = require("fs");
const { Packer, Document, Header, Footer, Paragraph } = require("docx");
const {
  header,
  footer,
  first_page_footer,
  first_page_header,
} = require("../utils/tables/header&footer");
const d = require("../utils/reportData");
const { styling } = require("../utils/styling");
const information = require("../utils/tables/information");
const summary = require("../utils/tables/summary");
const details = require("../utils/tables/details");
const othernote = require("../utils/tables/othernote");
const otherphoto = require("../utils/tables/otherphotos");
const sign = require("../utils/tables/sign");
const backcover = require("../utils/tables/backcover");
const { getCleanedString } = require("../utils/helper");
const export_word_path = "app/public/words";

class AppService extends Service {
  async generateWord() {
    try {
      const doc = new Document({
        styles: styling,
        sections: [
          {
            properties: {
              titlePage: true,
              page: {
                margin: {
                  top: 1000,
                  bottom: 1000,
                  left: 900, // 左右边距为900，则宽度为100%的表格，实际宽度为17.8cm
                  right: 900, // 左右边距为900，则宽度为100%的表格，实际宽度为17.8cm
                  header: 1000,
                  footer: 1000,
                },
              },
            },
            headers: {
              default: new Header({
                children: [header, new Paragraph("")],
              }),
              first: new Header({ children: [first_page_header] }),
            },
            footers: {
              default: new Footer({
                children: footer,
              }),
              first: new Footer({ children: first_page_footer }),
            },

            children: [
              new Paragraph(""),
              ...information,
              new Paragraph(""),
              ...summary,
              new Paragraph(""),
              ...details,
              new Paragraph(""),
              ...othernote,
              new Paragraph(""),
              ...otherphoto,
              new Paragraph(""),
              ...sign,
              new Paragraph({
                text: "",
                pageBreakBefore: true,
              }),
              ...backcover,
            ],
          },
        ],
      });

      const buf = await Packer.toBuffer(doc);
      const word_name = getCleanedString(d.ReportNo) + ".docx";
      fs.writeFileSync(`${export_word_path}/${word_name}`, buf);
      const host = "http://127.0.0.1:7001/";
      const word_path = host + "public/words/" + word_name;

      return { success: true, wordUrl: word_path };
    } catch (error) {
      console.log("Error:", error);
      return {
        success: false,
        message: error.message,
      };
    }
  }
}

module.exports = AppService;
