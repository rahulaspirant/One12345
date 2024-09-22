function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function doPost(form) {
  var sheet = SpreadsheetApp.openById('1R3D4cSKwSYRpSsbvwUsEDgzL7EgINi78UP6ZRBo0kI0').getSheetByName('Test'); // Replace with actual ID and name
  sheet.appendRow([form.date, form.projectNo, form.projectName, form.interfaces, form.activity, form.parameters, form.specification]);
}

////////////////////////////////////////////////////Index.html



<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <form id="myForm">
      <label for="date">Date:</label>
      <input type="date" id="date" name="date"><br><br>

      <label for="projectNo">Project No.:</label>
      <input type="text" id="projectNo" name="projectNo"><br><br>

      <label for="projectName">Project Name:</label>
      <input type="text" id="projectName" name="projectName"><br><br>

      <label for="interfaces">Interfaces:</label>
      <input type="text" id="interfaces" name="interfaces"><br><br>

      <table>
        <tr>
          <th>Activity</th>
          <th>Parameters</th>
          <th>Specification</th>
        </tr>
        <tr>
          <td><input type="text" name="activity"></td>
          <td><input type="text" name="parameters"></td>
          <td><input type="text" name="specification"></td>
        </tr>
      </table>

      <input type="submit" value="Submit">
    </form>

    <script>
      document.getElementById("myForm").addEventListener("submit", function(event) {
        event.preventDefault();
        var form = document.getElementById("myForm");
        google.script.run.doPost(form);
        form.reset();
      });
    </script>
  </body>
</html>

/////////////////////////////uploadfile .gs

Folder_Id = '1MuaUvDsNGQIVz9_PSVpbkFiZQeLCt1CE'


function onOpen(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var menuEntries = [];
  menuEntries.push({name: "File", functionName: "doGet"});
  ss.addMenu("Attach", menuEntries);
}


function upload(obj) {
  var file = DriveApp.getFolderById(Folder_Id).createFile(obj.upload);
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var File_name = file.getName()
  var value = 'hyperlink("' + file.getUrl() + '";"' + File_name + '")'
  
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  var cell = selection.getCurrentCell()
  cell.setFormula(value)
  
  return {
    fileId: file.getId(),
    mimeType: file.getMimeType(),
    fileName: file.getName(),
  };
}

function doGet(e) {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  var cell = selection.getCurrentCell();
  var html = HtmlService.createHtmlOutputFromFile('upload');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload File');
}


//////////////////////////////////////////////////////////// upload.html

<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
  </head>
  <body>
  <form> <!-- Modified -->
    <div id="progress" ></div>
    <input type="file" name="upload" id="file">
    <input type="button" value="Submit" class="action" onclick="form_data(this.parentNode)" >
    <input type="button" value="Close" onclick="google.script.host.close()" />
  </form>
  <script>
    function form_data(obj){ // Modified
      google.script.run.withSuccessHandler(closeIt).upload(obj);
    };
    function closeIt(e){ // Modified
      console.log(e);
      google.script.host.close();
    };
  </script>
</body>
</html>


