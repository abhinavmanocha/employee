window.addEventListener('load', function() {
    brython({
      debug: true
    });
  
    const excelFileInput = document.getElementById('excelFile');
    excelFileInput.addEventListener('change', handleFileSelect);
  });
  
  function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
      // Pass the file to the Python code
      sendFileToApp(file);
    }
  }
  
  function sendFileToApp(file) {
    // You can use the File API or other libraries to read the file contents
    // and pass it to the Python code
    // For example, using the FileReader API:
    const reader = new FileReader();
    reader.onload = function() {
      const fileContents = reader.result;
      // Call a Python function with the file contents
      runPythonCode(fileContents);
    };
    reader.readAsArrayBuffer(file);
  }
  
  function runPythonCode(fileContents) {
    // Call a Python function with the file contents
    // For example:
    pyApp.processExcelFile(fileContents);
  }
  