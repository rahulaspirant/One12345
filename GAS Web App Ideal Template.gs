function doGet(){
return HtmlService.createHtmlOutputFromFile('index');
}


function saveFormData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Responses");

  sheet.appendRow([
    data.product,
    data.description,
    data.quantity,
    //new Date()               
  ]);

  return "Saved";
}

function getData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Responses");
  const values = sheet.getDataRange().getValues();
  return values;

}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////

<!DOCTYPE html>
<html>
<head><base target="_top"></head>
<body>
  
  <!-- UI -->
  <form id="myform">
    <input id="product" placeholder="Product" required><br>
    <input id="description" placeholder="Description"><br>
    <input id="quantity" type="number" placeholder="Quantity"><br>
    <button id="submitBtn" type="submit">Submit</button>
  </form>

  <div id="status"></div>
  <div id="tableArea"></div>

  <script>
    // --------------------------
    // Helpers
    // --------------------------
    function showStatus(text) {
      document.getElementById('status').textContent = text || '';
    }

    function renderTable(values) {
      const container = document.getElementById('tableArea');
      if (!values || values.length === 0) {
        container.innerHTML = '<p>No data</p>';
        return;
      }

      let html = '<table border="1" cellpadding="5"><thead><tr>';
      values[0].forEach(h => html += `<th>${h}</th>`);
      html += '</tr></thead><tbody>';

      for (let i = 1; i < values.length; i++) {
        html += '<tr>';
        values[i].forEach(cell => html += `<td>${cell || '&nbsp;'}</td>`);
        html += '</tr>';
      }

      html += '</tbody></table>';
      container.innerHTML = html;
    }

    function getFormData() {
      return {
        product: document.getElementById('product').value.trim(),
        description: document.getElementById('description').value.trim(),
        quantity: document.getElementById('quantity').value
      };
    }

    // --------------------------
    // API wrappers
    // --------------------------
    function apiGetData(onSuccess, onFailure) {
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onFailure)
        .getData();
    }

    function apiSaveData(data, onSuccess, onFailure) {
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onFailure)
        .saveFormData(data);
    }

    // --------------------------
    // App Init
    // --------------------------
    function init() {
      const form = document.getElementById('myform');
      const submitBtn = document.getElementById('submitBtn');

      // Load existing data
      loadAndRenderData();

      form.addEventListener('submit', function (e) {
        e.preventDefault();
        const data = getFormData();

        if (!data.product) {
          showStatus('Product is required');
          return;
        }

        submitBtn.disabled = true;
        showStatus('Saving...');

        apiSaveData(
          data,
          function (resp) {
            showStatus(resp);
            form.reset();
            loadAndRenderData();
            submitBtn.disabled = false;
          },
          function (err) {
            showStatus('Error: ' + err.message);
            submitBtn.disabled = false;
          }
        );
      });
    }

    function loadAndRenderData() {
      apiGetData(renderTable, err => showStatus('Load error: ' + err.message));
    }

    document.addEventListener('DOMContentLoaded', init);
  </script>

</body>
</html>


