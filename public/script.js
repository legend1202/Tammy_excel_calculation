/* const url = "http://localhost:5000"; */
const url = "https://tammy-excel-calculation.onrender.com";

const nameBtn = document.getElementById("name_btn");
const fileBtn = document.getElementById("file_btn");

document
  .getElementById("upload-form")
  .addEventListener("submit", async function (event) {
    event.preventDefault();

    fileBtn.innerText = "Processing...";

    const fileInput = document.getElementById("file");
    if (!fileInput.files.length) {
      alert("Please select a file to upload.");
      return;
    }

    const formData = new FormData();
    formData.append("file", fileInput.files[0]);

    // Send the file to the backend for processing
    const response = await fetch(`${url}/upload-excel`, {
      method: "POST",
      body: formData,
    });
    if (response.ok) {
      const result = await response.json();
      const excelUrl = result.result.excelUrl;

      // Show download link
      const downloadSection = document.getElementById("download-section");
      const downloadLink = document.getElementById("download-link");
      downloadLink.href = url + excelUrl;
      downloadSection.style.display = "block";
    } else {
      alert("Error processing the file.");
    }

    fileBtn.innerText = "Upload and Process";
  });
