﻿function s() {
  Alert("123");
  return false;
}
var wrapper = document.getElementById("簽名板");
var canvas = wrapper.querySelector("canvas");
var savetodbButton = document.getElementById("save_to_db");
var signaturePad = new SignaturePad(canvas, {
  // It's Necessary to use an opaque color when saving image as JPEG;
  // this option can be omitted If only saving as PNG or SVG
  minWidth: 15,
  maxWidth: 20,
  backgroundColor: 'rgb(255, 255, 255)'
});
//var on_14= System.Configuration.ConfigurationManager["ApplicationServices"];
//alert(on_14);
// Adjust canvas coordinate space taking into account pixel ratio,
// to make it look crisp on mobile devices.
// This also causes canvas to be cleared.
function resizeCanvas() {
  // When zoomed out to less than 100%, for some very strange reason,
  // some browsers report devicePixelRatio as less than 1
  // and only part of the canvas is cleared then.
  var ratio = Math.max(window.devicePixelRatio || 1, 1);
  // This part causes the canvas to be cleared
  canvas.width = canvas.offsetWidth * ratio;
  canvas.height = canvas.offsetHeight * ratio;
  canvas.getContext("2d").scale(ratio, ratio);
  // This library does not listen for canvas changes, so after the canvas is automatically
  // cleared by the browser, SignaturePad#isEmpty might still return false, even though the
  // canvas looks empty, because the internal data of this library wasn't cleared. To make sure
  // that the state of this library is consistent with visual state of the canvas, you
  // have to clear it manually.
  signaturePad.clear();
}
// On mobile devices it might make more sense to listen to orientation change,
// rather than window resize events.
window.onresize = resizeCanvas;
resizeCanvas();
function download(dataURL, filename) {
  var blob = dataURLToBlob(dataURL);
  var url = window.URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.style = "display: none";
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
}
// One could simply use Canvas#toBlob method instead, but it's just to show
// that it can be done using result of SignaturePad#toDataURL.
function dataURLToBlob(dataURL) {
  // Code taken from https://github.com/ebidel/filer.js
  var parts = dataURL.split(';base64,');
  var contentType = parts[0].split(":")[1];
  var raw = window.atob(parts[1]);
  var rawLength = raw.length;
  var uInt8Array = new Uint8Array(rawLength);
  for (var i = 0; i < rawLength; ++i) {
    uInt8Array[i] = raw.charCodeAt(i);
  }
  return new Blob([uInt8Array], { type: contentType });
}
//簽名後按OK
savetodbButton.addEventListener("click", function (event) {
 //alert(Document.querySelectorAll())
 if (signaturePad.isEmpty()) {
    //alert("請簽名。");
    document.getElementById("簽名url").value = "";
    //document.getElementById("ContentPlaceHolder1_Button1").click();
  } else {
    var dataURL = signaturePad.toDataURL();
    document.getElementById("簽名url").value = dataURL;
        //document.getElementById("ContentPlaceHolder1_Button1").click();
  }
});
