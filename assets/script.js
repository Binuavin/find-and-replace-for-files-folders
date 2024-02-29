// Function to mimic placeholder for textarea
var textarea = document.getElementById("folderPath");
var placeholderText =
  "Paste the folder link here or click the [Select Folder] button..";

textarea.value = placeholderText;
textarea.style.color = "gray";

textarea.onfocus = function () {
  if (this.value === placeholderText) {
    this.value = "";
    this.style.color = "black";
  }
};

textarea.onblur = function () {
  if (this.value === "") {
    this.value = placeholderText;
    this.style.color = "gray";
  }
};

// Function to adjust window Size according to Page Content
function adjustWindowSize() {
  // var bodyWidth = document.body.scrollWidth;
  var bodyWidth = document.body.offsetWidth;
  // var bodyHeight = document.body.scrollHeight;
  var bodyHeight = document.body.offsetHeight;

  // alert(bodyHeight);
  if (bodyHeight < 950) {
    bodyHeight = 950;
  }

  window.resizeTo(bodyWidth, bodyHeight);
}

// Adjust window size when the page is loaded
window.onload = function () {
  adjustWindowSize();
};

// Open URL in Browser
function openLink(url) {
  var shell = new ActiveXObject("WScript.Shell");
  shell.Run(url);
}
