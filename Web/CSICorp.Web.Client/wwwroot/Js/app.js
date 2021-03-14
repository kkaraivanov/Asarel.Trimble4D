//window.saveAsFile = function (fileName, byteBase64) {
//    var link = this.document.createElement('a');
//    link.download = fileName;
//    link.href = "data:application/octet-stream;base64," + byteBase64;
//    this.document.body.appendChild(link);
//    link.click();
//    this.document.body.removeChild(link);
//}

function saveAsFile(filename, bytesBase64) {
    var link = document.createElement('a');
    link.download = filename;
    link.href = "data:application/octet-stream;base64," + bytesBase64;
    document.body.appendChild(link); // Needed for Firefox
    link.click();
    document.body.removeChild(link);
}