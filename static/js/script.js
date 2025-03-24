document.addEventListener("DOMContentLoaded", function () {
  // אלמנטים
  const dropArea = document.getElementById("drop-area");
  const fileInput = document.getElementById("file-input");
  const fileInfo = document.getElementById("file-info");
  const fileName = document.getElementById("file-name");
  const changeFileBtn = document.getElementById("change-file");
  const continueBtn = document.getElementById("continue-btn");
  const uploadSection = document.getElementById("upload-section");
  const fontSection = document.getElementById("font-section");
  const downloadSection = document.getElementById("download-section");
  const loadingSection = document.getElementById("loading-section");
  const backToUploadBtn = document.getElementById("back-to-upload");
  const processBtn = document.getElementById("process-btn");
  const fontSelect = document.getElementById("font-select");
  const fontPreviewText = document.getElementById("font-preview-text");
  const downloadBtn = document.getElementById("download-btn");
  const startOverBtn = document.getElementById("start-over");
  const selectedFontSpan = document.getElementById("selected-font");
  const step1 = document.getElementById("step1");
  const step2 = document.getElementById("step2");
  const step3 = document.getElementById("step3");

  // מידע קבצים
  let fileData = {
    id: null,
    originalName: null,
    selectedFont: null,
  };

  // טעינת רשימת הפונטים
  fetchFonts();

  // מעבר בין שלבים
  continueBtn.addEventListener("click", function () {
    uploadSection.style.display = "none";
    fontSection.style.display = "block";
    step1.classList.add("completed");
    step2.classList.add("active");
  });

  backToUploadBtn.addEventListener("click", function () {
    fontSection.style.display = "none";
    uploadSection.style.display = "block";
    step2.classList.remove("active");
    step1.classList.remove("completed");
  });

  startOverBtn.addEventListener("click", function () {
    downloadSection.style.display = "none";
    uploadSection.style.display = "block";
    resetForm();
    step3.classList.remove("active");
    step2.classList.remove("completed");
    step1.classList.remove("completed");
  });

  // גרירת קבצים
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropArea.addEventListener(eventName, highlight, false);
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, unhighlight, false);
  });

  function highlight() {
    dropArea.classList.add("drag-over");
  }

  function unhighlight() {
    dropArea.classList.remove("drag-over");
  }

  dropArea.addEventListener("drop", handleDrop, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;

    if (files.length) {
      fileInput.files = files;
      handleFiles(files);
    }
  }

  // ניהול בחירת קבצים
  fileInput.addEventListener("change", function () {
    if (fileInput.files.length) {
      handleFiles(fileInput.files);
    }
  });

  changeFileBtn.addEventListener("click", function () {
    fileInput.value = "";
    fileInfo.style.display = "none";
    document.querySelector(".upload-form").style.display = "block";
    continueBtn.disabled = true;
  });

  function handleFiles(files) {
    const file = files[0];

    if (file.name.toLowerCase().endsWith(".pptx")) {
      uploadFile(file);
    } else {
      alert("אנא בחרו קובץ PowerPoint (.pptx) תקין");
    }
  }

  // העלאת קובץ לשרת
  function uploadFile(file) {
    document.querySelector(".upload-form").style.display = "none";
    fileInfo.style.display = "flex";
    fileName.textContent = file.name;

    const formData = new FormData();
    formData.append("file", file);

    fetch("/upload", {
      method: "POST",
      body: formData,
    })
      .then((response) => response.json())
      .then((data) => {
        if (data.success) {
          fileData.id = data.file_id;
          fileData.originalName = data.original_name;
          continueBtn.disabled = false;
        } else {
          alert(data.error || "שגיאה בהעלאת הקובץ");
          resetForm();
        }
      })
      .catch((error) => {
        console.error("Error:", error);
        alert("שגיאה בהעלאת הקובץ");
        resetForm();
      });
  }

  // שינוי תצוגה מקדימה של פונט
  fontSelect.addEventListener("change", function () {
    fontPreviewText.style.fontFamily = this.value;
  });

  // עיבוד הקובץ
  processBtn.addEventListener("click", function () {
    const selectedFont = fontSelect.value;

    if (!selectedFont) {
      alert("אנא בחרו פונט");
      return;
    }

    fileData.selectedFont = selectedFont;
    selectedFontSpan.textContent = selectedFont;

    fontSection.style.display = "none";
    loadingSection.style.display = "block";

    fetch("/process", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        file_id: fileData.id,
        font: selectedFont,
        original_name: fileData.originalName,
      }),
    })
      .then((response) => response.json())
      .then((data) => {
        loadingSection.style.display = "none";

        if (data.success) {
          downloadSection.style.display = "block";
          downloadBtn.href = `/download/${fileData.id}/${fileData.originalName}`;
          step2.classList.add("completed");
          step3.classList.add("active");
        } else {
          alert(data.error || "שגיאה בעיבוד הקובץ");
          fontSection.style.display = "block";
        }
      })
      .catch((error) => {
        console.error("Error:", error);
        loadingSection.style.display = "none";
        fontSection.style.display = "block";
        alert("שגיאה בעיבוד הקובץ");
      });
  });

  // איפוס טופס
  function resetForm() {
    fileInput.value = "";
    fileInfo.style.display = "none";
    document.querySelector(".upload-form").style.display = "block";
    continueBtn.disabled = true;
    fileData = {
      id: null,
      originalName: null,
      selectedFont: null,
    };
  }

  // טעינת רשימת פונטים
  function fetchFonts() {
    fetch("/fonts")
      .then((response) => response.json())
      .then((fonts) => {
        fontSelect.innerHTML = "";

        // הוספת אפשרות ברירת מחדל
        const defaultOption = document.createElement("option");
        defaultOption.value = "";
        defaultOption.textContent = "בחרו פונט";
        defaultOption.disabled = true;
        defaultOption.selected = true;
        fontSelect.appendChild(defaultOption);

        // הוספת הפונטים מהשרת
        fonts.forEach((font) => {
          const option = document.createElement("option");
          option.value = font;
          option.textContent = font;
          option.style.fontFamily = font;
          fontSelect.appendChild(option);
        });
      })
      .catch((error) => {
        console.error("Error fetching fonts:", error);
        fontSelect.innerHTML =
          '<option value="" disabled selected>שגיאה בטעינת פונטים</option>';
      });
  }
});
