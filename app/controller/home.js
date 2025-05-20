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
        <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      line-height: 1.6;
      margin: 20px;
      background-color: #f4f4f4;
      color: #333;
      text-align: center;
    }

    .container {
      max-width: 800px; /* Optional: Limit the width of the content */
      margin: 0 auto; /* Center the container horizontally */
      background-color: #fff;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    p {
      margin-bottom: 1em;
      color: #555;
    }

    ol {
      list-style-type: decimal;
      padding-left: 20px;
      margin-bottom: 1.5em;
      color: #666;
    }

    li {
      margin-bottom: 0.5em;
      text-align: left;
    }

    .modern-button {
      display: inline-block;
      padding: 12px 24px;
      font-size: 16px;
      font-weight: bold;
      text-align: center;
      text-decoration: none;
      color: #fff;
      background-color: #007bff; /* A nice blue */
      border: none;
      border-radius: 6px;
      cursor: pointer;
      transition: background-color 0.3s ease;
    }

    .modern-button:hover {
      background-color: #0056b3;
    }

    .container {
      background-color: #fff;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
  </style>
      </head>
      <body>
      <div class="container">
        <h1>HQTS导出Word项目</h1>
        <p>以下是已确认的项目需求</p>
     <ol>
      <li>生成的报告文档需符合所提供的报告模板的整体布局与结构</li>
      <li>生成的报告文档内容需使用提供的SON格式报告数据结合报告模板按定逻辑动态生成报告</li>
      <li>生成报告文档的表格能够根据给定的业务规则进行表格的删除、新增、填充、排版等适配</li>
    </ol>

    <a class="modern-button" href="http://localhost:7001/api/generate-word/">点击这里验收成果</a>
    </div>
      </body>
      </html>
    `;
    ctx.type = "text/html";
    ctx.body = htmlContent;
  }
}

module.exports = HomeController;
