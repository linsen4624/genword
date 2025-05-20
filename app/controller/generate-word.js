const { Controller } = require("egg");
const { Packer, Document, Header, Footer, Paragraph } = require("docx");
const {
  header,
  footer,
  first_page_footer,
  first_page_header,
} = require("../utils/tables/header&footer");
const { styling } = require("../utils/styling");
const getInfoTable = require("../utils/tables/information");
const geSummaryTable = require("../utils/tables/summary");
const details = require("../utils/tables/details");
const othernote = require("../utils/tables/othernote");
const otherphoto = require("../utils/tables/otherphotos");
const sign = require("../utils/tables/sign");
const backcover = require("../utils/tables/backcover");

class GenerateWordController extends Controller {
  async create() {
    const { ctx } = this;

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
                  left: 980,
                  right: 980,
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
              getInfoTable(),
              new Paragraph(""),
              geSummaryTable(),
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

      ctx.set(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );
      ctx.set(
        "Content-Disposition",
        "attachment; filename=generated-document.docx"
      );

      ctx.body = buf;
    } catch (error) {
      console.log("Error:", error);
      ctx.body = {
        success: false,
        message: error.message,
      };
    }
  }
}

module.exports = GenerateWordController;
