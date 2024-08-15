    function ensureFullScreen() {
      const elem = document.documentElement;
      if (elem.requestFullscreen) {
        elem.requestFullscreen();
      } else if (elem.webkitRequestFullscreen) {
        elem.webkitRequestFullscreen();   

      } else if (elem.msRequestFullscreen) {
        elem.msRequestFullscreen();
      }
    }
	// Call the function when the window loads
	window.onload = ensureFullScreen;
	
function openExcelFile(fileKey) {
    eel.open_excel_file(fileKey)(function(result) {
        if (result) {
            console.log("Excel file opened successfully");
        } else {
            console.log("Failed to open Excel file");
        }
    });
}
