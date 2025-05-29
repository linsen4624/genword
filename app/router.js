/**
 * @param {Egg.Application} app - egg application
 */
module.exports = (app) => {
  const { router, controller } = app;
  router.get("/", controller.home.index);

  router.get("/api/generate-word", controller.generateWord.create);
  router.post("/api/generate-report", controller.apps.process_report);
  router.post("/uploadJSON", controller.upload.upload_json);
  router.post("/uploadZip", controller.upload.upload_zip);
};
