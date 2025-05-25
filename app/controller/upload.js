const { Controller } = require("egg");
const fs = require("fs");
const path = require("path");
const pump = require("mz-modules/pump");
const AdmZip = require("adm-zip");

const uploadPath = "app/public/upload/";
const json_target_path = "app/utils/";
const upzip_target_path = "images/";

class UploadController extends Controller {
  async upload_json() {
    const { ctx } = this;
    const stream = await ctx.getFileStream();

    if (!fs.existsSync(json_target_path)) fs.mkdirSync(json_target_path);

    let filename = encodeURIComponent(stream.filename);
    const extname = path.extname(filename);
    filename = "reportData" + extname;
    const target = path.join(this.config.baseDir, json_target_path, filename);
    const writeStream = fs.createWriteStream(target);
    await pump(stream, writeStream);
    ctx.body = { code: 0, message: "Successfully uploaded" };
  }

  async upload_zip() {
    const { ctx } = this;
    const stream = await ctx.getFileStream();
    if (!fs.existsSync(uploadPath)) fs.mkdirSync(uploadPath);
    let filename = encodeURIComponent(stream.filename);
    const extname = path.extname(filename);
    filename = "photo" + extname;
    const target = path.join(this.config.baseDir, uploadPath, filename);
    const writeStream = fs.createWriteStream(target);
    await pump(stream, writeStream);

    try {
      const zip = new AdmZip(uploadPath + filename);
      zip.extractAllTo(upzip_target_path, true);
    } catch (error) {
      console.error(`Error unzipping file "${upzip_target_path}":`, error);
    }

    ctx.body = { code: 0, message: "Successfully uploaded" };
  }
}

module.exports = UploadController;
