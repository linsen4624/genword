/* eslint valid-jsdoc: "off" */

/**
 * @param {Egg.EggAppInfo} appInfo app info
 */
module.exports = (appInfo) => {
  /**
   * built-in config
   * @type {Egg.EggAppConfig}
   **/
  const config = (exports = {});

  // use for cookie sign key, should change to your own and keep security
  config.keys = appInfo.name + "_1747193406802_1199";

  // add your middleware config here
  config.middleware = [];

  // add your user config here
  const userConfig = {
    // myAppName: 'egg',
  };

  config.multipart = {
    fileSize: "50mb",
    fileExtensions: [".json", ".zip"],
  };

  config.bodyParser = {
    enable: true,
    encoding: "utf8",
    formLimit: "5mb",
    jsonLimit: "5mb",
  };

  config.security = {
    csrf: {
      enable: false,
      ignoreJSON: true,
      headerName: "x-csrf-token",
      cookieName: "csrfToken",
    },
  };

  return {
    ...config,
    ...userConfig,
  };
};
