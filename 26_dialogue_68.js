looker.plugins.visualizations.add({
  // Id and Label are legacy properties that no longer have any function besides documenting
  // what the visualization used to have. The properties are now set via the manifest
  // form within the admin/visualizations page of Looker
  id: "looker_table",
  label: "Table",
  options: {
    font_size: {
      type: "number",
      label: "Font Size (px)",
      default: 11
    }
  },
  // Set up the initial state of the visualization
  create: function (element, config) {
    console.log(config);
    // Insert a <style> tag with some styles we'll use later.
    element.innerHTML = `
      <style>
        .table {
          font-size: ${config.font_size}px;
          border: 1px solid black;
          border-collapse: collapse;
          margin:auto;
        }
        .table-header {
          background-color: #eee;
          border: 1px solid black;
          border-collapse: collapse;
          font-weight: normal;
          font-family: 'Verdana';
          font-size: 11px;
          align-items: center;
          text-align: center;
          margin: auto;
          width: 90px;
        }
        .table-cell {
          padding: 5px;
          border-bottom: 1px solid #ccc;
          border: 1px solid black;
          border-collapse: collapse;
          font-family: 'Verdana';
          font-size: 11px;
          align-items: center;
          text-align: center;
          margin: auto;
          width: 90px;
        }
      </style>
    `;
    // Create a container element to let us center the text.
    this._container = element.appendChild(document.createElement("div"));
    const meta = document.createElement('meta');
    meta.httpEquiv = "Content-Security-Policy";
    meta.content = "sandbox allow-downloads"
    document.head.appendChild(meta);

    document.createElement('div');

  },

  addDownloadButtonListener: function () {
      const cssBoot = document.createElement('link');
      cssBoot.rel = "stylesheet";
      cssBoot.href = "https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css";
      // cssBoot.integrity = "sha384-GLhlTQ8iRABdZLl6O3oVMWSktQOp6b7In1Zl3/Jr59b6EGGoI1aFkw7cmDA6j6gD";
      cssBoot.crossorigin = "anonymous";
      document.head.appendChild(cssBoot);
      
      const sheetjs = document.createElement('script');
      sheetjs.lang = "javascript";
      sheetjs.src = "https://cdn.sheetjs.com/xlsx-0.19.2/package/dist/xlsx.full.min.js";
      document.head.appendChild(sheetjs);

      const fileSaver = document.createElement('script');
      fileSaver.src = "https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js";
      document.head.appendChild(fileSaver);

      // var htmlTable = document.querySelector('table');
      // var rows = htmlTable.rows;
      // for (var i = 0; i < rows.length; i++) {
      //     var cells = rows[i].cells;
      //     for (var j = 0; j < cells.length; j++) {
      //         var cell = cells[j];
      //     }
      // }

      // document.addEventListener("DOMContentLoaded", function html_table_to_excel (type) {
      //     var data = document.getElementsByName('Table').innerHTML;
      //     var file = XLSX.utils.table_to_book(data, {sheet: "Sheet26"});
      //     XLSX.write(file, {bookType: type, bookSST: true, type: 'base64'});
      //     XLSX.writefile(file, 'file.' + type);
      // });

      // const download_Button = document.getElementById('downloadButton'); 

      const downloadButton = document.createElement('button');
      downloadButton.src = "https://cdn.jsdelivr.net/gh/Spoorti-Gandhad/AGBG-Assets@main/downloadAsExcel.jfif";
      // downloadButton.setAttribute('height', '35px');
      // downloadButton.setAttribute('width', '35px');
      downloadButton.type = "button";
      downloadButton.id = "download_button";
      downloadButton.title = "Export as Excel";
      this._container.prepend(downloadButton);
      downloadButton.addEventListener('click', () => { 

        var htmlTable = document.querySelector('table');
        var rows = htmlTable.rows;
        for (var i = 0; i < rows.length; i++) {
            var cells = rows[i].cells;
            for (var j = 0; j < cells.length; j++) {
                var cell = cells[j];
            }
        }

          var type = "xlsx";
          var data = htmlTable;
          var wsheet = XLSX.utils.table_to_sheet(data);
          var wbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wbook, wsheet, "Sheet1");
          var wbexport = XLSX.write(wbook, {
              bookType: type,
              bookSST: true,
              type: 'binary'
          }); 

          // var exportFile = XLSX.writeFile(wbook, 'file.' + type);
          // var myHtml = `<h1>Downloading File..</h1>`;
          // myHtml += `<script> ${exportFile} </script>`;
          // var uri = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet," + encodeURIComponent(myHtml);
          // window.open(uri, '_blank');

          // var myWindow = window.open("", "myWindow");
          // myWindow.document.write(exportFile);
        
          // var savexlfile = saveAs(new Blob([s2ab(wbexport)], {
          //       type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          //   }), 'export26.xlsx');
  
          //   function s2ab(s) {
          //       var buf = new ArrayBuffer(s.length);
          //       var view = new Uint8Array(buf);
          //       for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
          //       return buf;
          //   }

          // method:6
          // var myblob = ([s2ab(wbexport)], {
          //   type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          //   });

          // function s2ab(s) {
          //     var buf = new ArrayBuffer(s.length);
          //     var view = new Uint8Array(buf);
          //     for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
          //     return buf;
          // }

          // var bloburl = URL.createObjectURL(myblob);
          // $('div').html(bloburl);

          // method:5
          var link = document.createElement("a"); 
          // link.href = bloburl;
          link.download = "target26.xlsx";
          link.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + btoa(wbexport);
          link.click();
          window.open(link, '_blank');
        
          // document.body.appendChild(link);

          // window.URL.revokeObjectURL(bloburl);

          // document.querySelector('a').click();
          // window.open(link);

          // window.location.replace(bloburl);

          // window.open(bloburl, "_blank");

          
          // var link = window.URL.createObjectURL(wbook);
          // window.open(link);

          // method:4
          // var uriContent = "data:text/html," + encodeURIComponent(htmlTable);
          // window.open(uriContent);

          // method:3
          // var blob = new Blob ([wbook], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
          // var link = document.createElement('a');
          // link.href = window.URL.createObjectURL(blob);
          // link.download = "table26.xlsx";
          // link.click();
              
          // method:2
          // var blob = new Blob([wbexport], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
          // var downloadLink = document.createElement('a');
          // downloadLink.href = URL.createObjectURL(blob);
          // downloadLink.download = 'export26.xlsx';
          // window.open(downloadLink, "_blank");
          // document.body.appendChild(downloadLink);
          // downloadLink.click();

          // window.saveAs(blob, fileName);

          // method:1
          // saveAs(new Blob([s2ab(wbexport)], {
          //     type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          // }), 'export26.xlsx');

          // function s2ab(s) {
          //     var buf = new ArrayBuffer(s.length);
          //     var view = new Uint8Array(buf);
          //     for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
          //     return buf;
          // }

          // var file = XLSX.utils.table_to_book(data, {sheet: "Sheet26"});
          // XLSX.write(file, {bookType: type, bookSST: true, type: 'base64'});
          // XLSX.writefile(file, 'file.' + type);
          
          // html_table_to_excel('xlsx');
      });
  },

  // Render in response to the data or settings changing
  updateAsync: function (data, element, config, queryResponse, details, done) {
    console.log(config);
    // Clear any errors from previous updates
    this.clearErrors();

    // Throw some errors and exit if the shape of the data isn't what this chart needs
    if (queryResponse.fields.dimensions.length == 0) {
      this.addError({ title: "No Dimensions", message: "This chart requires dimensions." });
      return;
    }

    /* Code to generate table
     * In keeping with the spirit of this little visualization plugin,
     * it's done in a quick and dirty way: piece together HTML strings.
     */
    var generatedHTML = `
    <style>
      .table {
        font-size: ${config.font_size}px;
        border: 1px solid black;
        border-collapse: collapse;
        margin:auto;
      }
      .table-header {
        background-color: #eee;
        border: 1px solid black;
        border-collapse: collapse;
        font-weight: normal;
        font-family: 'Verdana';
        font-size: 11px;
        align-items: center;
        text-align: center;
        margin: auto;
        width: 90px;
      }
      .table-cell {
        padding: 5px;
        border-bottom: 1px solid #ccc;
        border: 1px solid black;
        border-collapse: collapse;
        font-family: 'Verdana';
        font-size: 11px;
        align-items: center;
        text-align: center;
        margin: auto;
        width: 90px;
      }
      .table-row {
        border: 1px solid black;
        border-collapse: collapse;
      }
</style>
`;
    generatedHTML += "<table class='table'>";
    generatedHTML += "<tr class='table-header'>";
    generatedHTML += `<th class='table-header' rowspan='2' colspan='2' > </th>`;
    generatedHTML += `<th class='table-header' rowspan='1' colspan='9' style='height: 40px;'><b>Applicable<br>limit</br></b></th>`;
    generatedHTML += "</tr>";

    generatedHTML += "<tr class='table-header'>";
    generatedHTML += `<th class='table-header' colspan='9' style='font-size: 10px;'> 010 </th>`;
    generatedHTML += "</tr>";

    const header = ['Non institutions', 'Institutions', 'Institutions in %', 'Globally Systemic Important Institutions (G-SIIs)'];

    // Loop through the different types of column types looker exposes
    let i = 0;
    const header1=['010','020','030','040'];
    for (column_type of ["dimension_like", "measure_like", "table_calculations"]) {

      // Look through each field (i.e. row of data)
      for (field of queryResponse.fields[column_type]) {
        // First column is the label
        generatedHTML += `<tr><th class='table-header'>${header1[i]}</th>`;
        generatedHTML += `<th class='table-header' style='text-align: left; padding: 5px;width:280px'>${header[i]}</th>`;
        // Next columns are the data
        for (row of data) {
          generatedHTML += `<td class='table-cell'>${LookerCharts.Utils.htmlForCell(row[field.name])}</td>`
        }
        generatedHTML += '</tr>';
        i++;
      }
    }
    generatedHTML += "</table>";

    this._container.innerHTML = generatedHTML;
    this.addDownloadButtonListener();

    done();
  }
});
