const { Controller } = require("egg");
const fs = require("fs");
const path = require("path");
const pump = require("mz-modules/pump");
const AdmZip = require("adm-zip");
const json_target_path = "app/utils/";
const { upzip_target_path } = require("../utils/styling");
const { getCleanedString } = require("../utils/helper");

class AppController extends Controller {
  async process_report() {
    const { ctx } = this;
    const stream = await ctx.getFileStream();
    if (!fs.existsSync(json_target_path)) fs.mkdirSync(json_target_path);

    let filename = encodeURIComponent(stream.filename);
    const extname = path.extname(filename);
    filename = "reportData" + extname;
    const target = path.join(this.config.baseDir, json_target_path, filename);
    const writeStream = fs.createWriteStream(target);
    await pump(stream, writeStream);
    const d = fs.readFileSync(target);
    const data = JSON.parse(d);

    if (data.photosPackageURL) {
      try {
        const result = await ctx.curl(data.photosPackageURL, {
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
        zip.extractAllTo(
          `${upzip_target_path}/${getCleanedString(data.ReportNo)}`,
          true
        );
      } catch (e) {
        ctx.body = { result: "error", msg: "cannot handle zip file" };
      }

      ctx.body = await ctx.service.apps.generateWord();
    } else {
      ctx.body = { result: "error", msg: "no photo file found" };
    }
  }
}

module.exports = AppController;
