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
            postion: inherit;
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
          .thead{
            position: sticky;
            top: 0px; 
            z-index: 3;
          }
          th:after {
            content:''; 
            position:absolute; 
            left: 0; 
            bottom: 0; 
            width:100%; 
            border-bottom: 1px solid rgba(0,0,0,0.12);
          }
          th:before {
            left: 0;
            position: absolute;
            content: '';
            width: 100%;
            border-top: 1px solid #4c535b;
            top: 103px;
         }
         .div{
            overflow-y: auto;
            height: calc(100vh - 100px);
            margin-bottom: 100px;
            border-bottom: 0.5px solid black;
        }
        </style>
      `;
  
      // Create a container element to let us center the text.
      const div = document.createElement("div");
      div.classList.add('div');
      this._container = element.appendChild(div);
  
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
    
        // const fileSaver = document.createElement('script');
        // fileSaver.src = "https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js";
        // document.head.appendChild(fileSaver);

        const xlsxstyle = document.createElement('script');
        xlsxstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-style@0.8.13/dist/xlsx.full.min.js";
        document.head.appendChild(xlsxstyle);
    
        const xlsxjsstyle = document.createElement('script');
        // xlsxjsstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js";
        xlsxjsstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
        document.head.appendChild(xlsxjsstyle);
    
        const downloadButton = document.createElement('img');
        downloadButton.src = "https://cdn.jsdelivr.net/gh/Spoorti-Gandhad/AGBG-Assets@main/downloadAsExcel.jfif";
        downloadButton.setAttribute('height', '25px');
        downloadButton.setAttribute('width', '25px');
        downloadButton.setAttribute('title', 'Download As Excel'); 
         downloadButton.style.marginLeft='90%';
        // downloadButton.type = "button";
        // downloadButton.id = "download_button";
        // downloadButton.title = "Export as Excel";
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
            var tdata = htmlTable;
            var sheader = [{v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, fill: {patternType: "solid", fgColor: {rgb: "FFFFFF"}, bgColor: {rgb: "A9AAAB"}}, border: {bottom: {style: "medium"}, right: {style: "medium"}}}}];
            var note = [{v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 10}}}];
            var wsheet = XLSX.utils.table_to_sheet(tdata, {origin: 'A4'});
            wsheet["!merges"] = [{s:{c:0, r:0}, e:{c:7, r:0}}, {s:{c:0, r:1}, e:{c:7, r:1}}, {s:{c:0, r:3}, e:{c:7, r:3}}];
            XLSX.utils.sheet_add_aoa(wsheet, [sheader], { origin: 'A1' });
            XLSX.utils.sheet_add_aoa(wsheet, [note], { origin: 'A2' });
            var wbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wbook, wsheet, "C27");
            var wbexport = XLSX.write(wbook, {
                bookType: type,
                bookSST: true,
                type: 'binary',
                cellStyles: true
            }); 
            
            var link = document.createElement("a"); 
            link.download = "target26.xlsx";
            link.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + btoa(wbexport);
            link.click();
            window.open(link, '_blank');
          
        });
    },
  
    // addDownloadButtonListener: function () {
    // const downloadButton = document.createElement('img');
    // downloadButton.src = "https://cdn.jsdelivr.net/gh/Spoorti-Gandhad/AGBG-Assets@main/downloadAsExcel.jfif";
    // downloadButton.setAttribute('height', '25px');
    // downloadButton.setAttribute('width', '25px');
    // downloadButton.setAttribute('title', 'Download As Excel');   
    // downloadButton.style.marginLeft='90%';
    // this._container.prepend(downloadButton);
    // downloadButton.addEventListener('click', (event) => {
    //       var uri = 'data:application/vnd.ms-excel;base64,'
    //         , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{Worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body><table>{table}</table></body></html>'
    //         , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
    //         , format = function (s, c) {
    //           const regex = /style="([^"]*)"/g;
    //           return s.replace(/{(\w+)}/g, function (m, p) {
    //             const cellHtml = c[p];
    //             const cellHtmlWithStyle = cellHtml.replace(regex, function (m, p1) {
    //               return 'style="' + p1 + '"';
    //             });
    //             return cellHtmlWithStyle;
    //           });
    //         };
    //      // Create a new style element and set the default styles
    //     var table = document.querySelector('table');  
    //   // table.style.type = 'text/css';
    //   // table.style.innerHTML = 'td, th { background-color: white; border: 1px solid black; font-weight: normal; font-size: 11pt; font-family: Calibri; mso-number-format: "\\\@"; }';
         
    //   var rows = table.rows;
    //     for (var i = 0; i < rows.length; i++) {
    //     var cells = rows[i].cells;
    //     for (var j = 0; j < cells.length; j++) {
    //       var cell = cells[j];
          
    //     //   cell.setAttribute('style');
    //      }
    //     }
    //       const XLSX = document.createElement('script');
    //       XLSX.src = 'https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js';
    //       document.head.appendChild(XLSX);
    //       //var ctx = { Worksheet: '27', table: table.innerHTML };
    //       var ctx = { Worksheet: '27', table: "<tr class='table-header'><th class='table-header' rowspan='1' colspan='8' style='align-items: left;text-align: left; height: 40px;border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>C 27.00 - Identification of the counterparty (LE 1)</b></th></tr><tr class='table-header'><th class='table-header' rowspan='1' colspan='3' style='background-color:none !important;font-family:Verdana;font-size:10px;align-items: center;text-align: left;padding: 5px;color:grey;font-weight:normal;'>* All values reported are in millions </th></tr>"+table.innerHTML }
         
    //       var xl = format(template, ctx);
    //       const downloadUrl = uri + base64(xl);
    //       console.log(downloadUrl); // Prints the download URL to the console
    //       //sleep(1000);
    //       //window.open(downloadUrl);
    //       window.open(downloadUrl, "_blank");
    //       //setTimeout(window.open(downloadUrl, 'Download'),1000);
    //     });
    //   },
    
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
            postion: inherit;
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
          .thead{
            position: sticky;
            top: 0px; 
            z-index: 3;
          }
          th:after {
            content:''; 
            position:absolute; 
            left: 0; 
            bottom: 0; 
            width:100%; 
            border-bottom: 1px solid rgba(0,0,0,0.12);
         }
        th:before {
            left: 0;
            position: absolute;
            content: '';
            width: 100%;
            border-top: 1px solid #4c535b;
            top: 103px;
        }
        .div{
            overflow-y: auto;
            height: calc(100vh - 100px);
            margin-bottom: 100px;
            border-bottom: 0.5px solid black;
        }
        </style>
      `;
  
      generatedHTML += "<p style='font-family:Verdana;margin:auto;font-weight:bold;font-size:14px;align-items:center;text-align:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 27.00 - Identification of the counterparty (LE 1)</p>";
      generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;text-align: right;padding: 5px;'>* All values reported are in millions </p>";
      generatedHTML += "<table class='table'>";
      generatedHTML += "<thead class='thead'>";
      generatedHTML += "<tr class='table-header'>";
      generatedHTML += "<th class='table-header' colspan='8' style='font-weight: bold;height:19px;border: 1px solid black;background-color: #eee;font-family: Verdana;width: -webkit-fill-available; position: absolute;'>COUNTERPARTY IDENTIFICATION</th>";
      generatedHTML += "</tr>";
      generatedHTML += "<tr class='table-header'>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Code</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Type of Code</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Name</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>National Code</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Residence of the Counterparty</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Sector of the Counterparty</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>NACE Code</th>";
      generatedHTML += "<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;height:100px;'>Type of Counterparty</th>";
      generatedHTML += "</tr>";
     
      const header=['011','015','021','035','040','050','060','070'];
      // First row is the header
      generatedHTML += "<tr class='table-header'>";
      for (let i=0;i<header.length;i++) {
        generatedHTML += `<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;mso-number-format: "\ \@";'>${header[i]}</th>`;
      }
      generatedHTML += "</tr>";
      generatedHTML += "</thead>";
  
      // Next rows are the data
      for (row of data) {
        generatedHTML += "<tr class='table-row'>";
        for (field of queryResponse.fields.dimensions.concat(queryResponse.fields.measures)) {
          generatedHTML += `<td class='table-cell' style='border: 1px solid black;'>${LookerCharts.Utils.htmlForCell(row[field.name])}</td>`;
        }
        generatedHTML += "</tr>";
      }
      generatedHTML += "</table>";
  
      this._container.innerHTML = generatedHTML;
      this.addDownloadButtonListener();
      done();
      
    }
  });
