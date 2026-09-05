//Adding two times
function addTimes(time1, time2) {
  // Helper to convert HH:MM:SS to total seconds
  const toSeconds = (time) => {
    const [h, m, s] = time.split(":").map(Number);
    return h * 3600 + m * 60 + s;
  };

  // Add the total seconds together
  let totalSeconds = toSeconds(time1) + toSeconds(time2);

  // Extract hours, minutes, and seconds from the total
  const hours = Math.floor(totalSeconds / 3600);
  totalSeconds %= 3600;

  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;

  // Pad numbers with leading zeros to match HH:MM:SS format
  const pad = (num) => String(num).padStart(2, "0");

  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
}

//Date functional formating
const dateFormat = (date) => {
  const dateStr = date;

  const [datePart, timePart] = dateStr.split(" ");
  const [month, day, year] = datePart.split("/");
  const [hours, minutes, seconds] = timePart.split(":");

  const newDate = new Date(year, month - 1, day);

  return newDate;
};

//Change keys of object
const changeKeys = (item) => {
  let result = {};
  for (const [key, value] of Object.entries(item)) {
    const newKey = key.toLowerCase().replace(/\s+/g, "").split("(mb)")[0];
    result[newKey] =
      newKey === "logintime"
        ? dateFormat(value).toDateString()
        : newKey === "download" || newKey === "upload"
          ? parseFloat(value.trim()).toFixed(2)
          : value;
  }

  return result;
};

//Duplicate Date merge
const duplicateDateMerge = (array) => {
  let newArr = [];

  for (let i = 0; i < array.length; i++) {
    let objModify = changeKeys(array[i]);
    const findIndx = newArr.findIndex(
      (i) => i.logintime === objModify.logintime,
    );
    if (
      newArr.filter((i) => i.logintime === objModify.logintime).length === 0
    ) {
      newArr.push(objModify);
    }

    if (newArr.length > 0 && findIndx > 0) {
      newArr[findIndx].download = (
        Number(newArr[findIndx].download) + Number(objModify.download)
      ).toFixed(2);
      newArr[findIndx].upload = (
        Number(newArr[findIndx].upload) + Number(objModify.upload)
      ).toFixed(2);
      newArr[findIndx].sessiontime = addTimes(
        newArr[findIndx].sessiontime,
        objModify.sessiontime,
      );
    }
  }

  return newArr;
};

// Upload an Excel or CSV file and display its summary in the table.
function uploadFile() {
  const fileInput = document.getElementById("excelFile");
  const file = fileInput.files && fileInput.files[0];

  if (!file) {
    alert("Please select an Excel or CSV file first.");
    return;
  }

  const fileExtension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xls", "xlsx"].includes(fileExtension)) {
    alert("Please select a valid Excel or CSV file.");
    return;
  }

  const uploadButton = document.getElementById("uploadBtn");
  const originalButtonContent = uploadButton.innerHTML;
  uploadButton.disabled = true;
  uploadButton.innerHTML =
    '<span class="spinner-border spinner-border-sm me-2" aria-hidden="true"></span>Loading...';

  const finishUpload = () => {
    setTimeout(() => {
      uploadButton.disabled = false;
      uploadButton.innerHTML = originalButtonContent;
    }, 2000);
  };

  const reader = new FileReader();
  reader.onload = function (event) {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        defval: "",
        raw: false,
      });
      const uniqueData = duplicateDateMerge(jsonData);
      setTimeout(() => {
        document.querySelector(".data-summary-row").style.display = "flex";

        const tableBody = document.getElementById("dataTableBody");
        tableBody.innerHTML = "";
        uniqueData.forEach((item) => {
          const row = tableBody.insertRow();
          const dateCell = row.insertCell(0);
          const downloadCell = row.insertCell(1);
          const uploadCell = row.insertCell(2);
          const sessionTimeCell = row.insertCell(3);
          dateCell.textContent = item.logintime;
          downloadCell.textContent = item.download;
          uploadCell.textContent = item.upload;
          sessionTimeCell.textContent = item.sessiontime;
        });
      }, 2000);
    } catch (error) {
      alert("The selected file could not be processed.");
    } finally {
      finishUpload();
    }
  };

  reader.onerror = function () {
    uploadButton.disabled = false;
    uploadButton.innerHTML = originalButtonContent;
    alert("The selected file could not be read.");
  };

  reader.readAsArrayBuffer(file);
}

//Download summary data as Excel file
function downloadSummary() {
  const tableBody = document.getElementById("dataTableBody");

  if (!tableBody || tableBody.rows.length === 0) {
    alert("No data available to download.");
    return;
  }

  const downloadButton = document.getElementById("downloadSummaryBtn");
  const originalButtonContent = downloadButton.innerHTML;
  downloadButton.disabled = true;
  downloadButton.innerHTML =
    '<span class="spinner-border spinner-border-sm me-2" aria-hidden="true"></span>Loading...';

  const rows = [];
  const headers = ["Date", "Download", "Upload", "Session Time"];
  rows.push(headers);

  Array.from(tableBody.rows).forEach((row) => {
    const values = Array.from(row.cells).map((cell) => cell.textContent.trim());
    rows.push(values);
  });

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const headerStyle = { font: { bold: true } };

  for (let col = 0; col < headers.length; col++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
    const cell = worksheet[cellRef];

    worksheet[cellRef] = {
      ...cell,
      v: headers[col],
      t: "s",
      s: headerStyle,
    };
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");

  setTimeout(() => {
    XLSX.writeFile(workbook, "data-summary.xlsx");
    downloadButton.disabled = false;
    downloadButton.innerHTML = originalButtonContent;
  }, 2000);
}

function clearFile() {
  const fileInput = document.getElementById("excelFile");
  fileInput.value = "";
  const tableBody = document.getElementById("dataTableBody");
  tableBody.innerHTML = "";
  document.querySelector(".data-summary-row").style.display = "none";
}
