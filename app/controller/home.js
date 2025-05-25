const { Controller } = require("egg");

class HomeController extends Controller {
  async index() {
    const { ctx } = this;
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>HQTS导出Word项目</title>
        <link rel="stylesheet" href="/public/css/style.css">
      </head>
      <body>
      <div class="upload-container">
        <h1>HQTS导出Word项目</h1>
        <p>注意！上传后的JSON文件和解压缩后的图片文件夹将会直接放置在使用目录，数据和图片是这个项目能够正常运行的关键依赖项，上传前请确保数据完整！</p>
       
        <form id="uploadJSONForm" enctype="multipart/form-data">
        <label class="upload-box">
            <input type="file" id="jsonFileInput" name="jsonFile" accept=".json" required>
            <p id="uploadJSONText">Click here to upload file</p>
            <span id="jsonFileNameDisplay" class="file-name"></span>
            <button id="uploadJSONButton" class="upload-button" type="submit" style="display:none">Upload JSON Data</button>
            <div id="json_status"></div>
          </label>
        </form>

<p></p>
        
        <form id="uploadZIPForm" enctype="multipart/form-data">
        <label class="upload-box">
            <input type="file" id="zipFileInput" name="zipFile" accept=".zip" required>
            <p id="uploadZIPText">Click here to Zip file</p>
            <span id="zipFileNameDisplay" class="file-name"></span>
            <button id="uploadZIPButton" class="upload-button" type="submit" style="display:none">Upload Zip File</button>
            <div id="zip_status"></div>
            </label>
        </form>
        
<p></p>
<hr class="minimal"/>

        <a class="upload-button" href="http://localhost:7001/api/generate-word/">点击这里验收成果</a>
        </div>
      </body>
      <script src="/public/js/home.js"></script>
      </html>
    `;
    ctx.type = "text/html";
    ctx.body = htmlContent;
  }
}

module.exports = HomeController;
