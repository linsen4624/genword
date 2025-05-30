const { Controller } = require("egg");
const fs = require("fs");
const AdmZip = require("adm-zip");
const json_target_path = "app/utils/";
const { upzip_target_path } = require("../utils/styling");
const { getCleanedString } = require("../utils/helper");
const d = require("../utils/reportData.json");

class AppController extends Controller {
  async process_report() {
    const { ctx } = this;
    const json_data = ctx.request.body;
    const zip_file_path = json_data.photosPackageURL;
    const report_no = json_data.ReportNo;

    fs.writeFileSync(
      json_target_path + "reportData.json",
      JSON.stringify(json_data, null, 2),
      "utf8"
    );

    if (zip_file_path !== "") {
      try {
        const result = await ctx.curl(zip_file_path, {
          dataType: "buffer",
          timeout: 10000,
        });

        if (result.status !== 200) {
          throw new Error(
            `Failed to download zip file. HTTP Status: ${result.status}`
          );
        }
        const zipBuffer = result.data;

        if (zipBuffer.length === 0) {
          throw new Error("Downloaded zip file is empty.");
        }

        const zip = new AdmZip(zipBuffer);
        if (!fs.existsSync(upzip_target_path)) fs.mkdirSync(upzip_target_path);
        zip.extractAllTo(
          `${upzip_target_path}/${getCleanedString(report_no)}`,
          true
        );
      } catch (e) {
        ctx.body = { result: "error", msg: "cannot handle zip file" };
      }

      if (d && Object.keys(d).length > 10) {
        ctx.body = await ctx.service.apps.generateWord();
      } else {
        ctx.body = { result: "error", msg: "no data file found" };
      }
    } else {
      ctx.body = { result: "error", msg: "no photo file found" };
    }
  }
}

module.exports = AppController;
