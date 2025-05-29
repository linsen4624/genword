/* eslint-disable space-before-function-paren */
/* eslint-disable no-undef */

document.addEventListener("DOMContentLoaded", () => {
  const json_file_Input = document.getElementById("jsonFileInput");
  const zip_file_Input = document.getElementById("zipFileInput");
  const upload_json_text = document.getElementById("uploadJSONText");
  const upload_zip_text = document.getElementById("uploadZIPText");
  const upload_json_btn = document.getElementById("uploadJSONButton");
  const upload_zip_btn = document.getElementById("uploadZIPButton");
  const json_fileNameDisplay = document.getElementById("jsonFileNameDisplay");
  const zip_fileNameDisplay = document.getElementById("zipFileNameDisplay");
  const uploadBox = document.querySelector(".upload-box");

  const jsonForm = document.getElementById("uploadJSONForm");
  const zipForm = document.getElementById("uploadZIPForm");

  json_file_Input.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      const fileName = e.target.files[0].name;

      json_fileNameDisplay.textContent = fileName;
      json_fileNameDisplay.classList.add("visible");
      upload_json_btn.style.display = "inline";

      uploadBox.classList.add("has-file");
      upload_json_text.textContent = "File selected:";
    } else {
      json_fileNameDisplay.textContent = "";
      json_fileNameDisplay.classList.remove("visible");
      upload_json_btn.style.display = "none";
      uploadBox.classList.remove("has-file");
      upload_json_text.textContent = "Click here to upload files";
    }
  });

  zip_file_Input.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      const fileName = e.target.files[0].name;

      zip_fileNameDisplay.textContent = fileName;
      zip_fileNameDisplay.classList.add("visible");
      upload_zip_btn.style.display = "inline";

      uploadBox.classList.add("has-file");
      upload_zip_text.textContent = "File selected:";
    } else {
      zip_fileNameDisplay.textContent = "";
      zip_fileNameDisplay.classList.remove("visible");
      upload_zip_btn.style.display = "none";
      uploadBox.classList.remove("has-file");
      upload_zip_text.textContent = "Click here to upload files";
    }
  });

  jsonForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    const statusDiv = document.getElementById("json_status");
    statusDiv.textContent = "Uploading...";
    statusDiv.style.color = "blue";
    const formData = new FormData();
    const fileInput = this.querySelector('input[type="file"]');
    const file = fileInput.files[0];

    if (!file) {
      statusDiv.textContent = "Please select a file.";
      statusDiv.style.color = "red";
      return;
    }
    formData.append("jsonFile", file);

    try {
      const response = await fetch("/api/uploadJSON", {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (result.code === 0) {
        statusDiv.textContent = `Success: ${result.message}`;
        statusDiv.style.color = "green";
      } else {
        statusDiv.textContent = `Error: ${result.message}`;
        statusDiv.style.color = "red";
      }
    } catch (error) {
      statusDiv.textContent = `Network Error: ${error.message}`;
      statusDiv.style.color = "red";
      console.error("Upload failed:", error);
    }
  });

  zipForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    const statusDiv = document.getElementById("zip_status");
    statusDiv.textContent = "Uploading...";
    statusDiv.style.color = "blue";
    const formData = new FormData();
    const fileInput = this.querySelector('input[type="file"]');
    const file = fileInput.files[0];

    if (!file) {
      statusDiv.textContent = "Please select a file.";
      statusDiv.style.color = "red";
      return;
    }
    formData.append("zipFile", file);

    try {
      const response = await fetch("/uploadZip", {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (result.code === 0) {
        statusDiv.textContent = `Success: ${result.message}`;
        statusDiv.style.color = "green";
      } else {
        statusDiv.textContent = `Error: ${result.message}`;
        statusDiv.style.color = "red";
      }
    } catch (error) {
      statusDiv.textContent = `Network Error: ${error.message}`;
      statusDiv.style.color = "red";
      console.error("Upload failed:", error);
    }
  });
});
