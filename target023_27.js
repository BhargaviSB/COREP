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
        const meta = document.createElement('meta');
    meta.httpEquiv = 'cache-control';
    meta.content = 'no-cache, no-store, must-revalidate';
    document.head.appendChild(meta);
    const meta2 = document.createElement('meta');
    meta2.httpEquiv = 'expires';
    meta2.content = '0';
    document.head.appendChild(meta2);
    const meta3 = document.createElement('meta');
    meta3.httpEquiv = 'pragma';
    meta3.content = 'no-cache';
    document.head.appendChild(meta3);
  
    },
  
    addDownloadButtonListener: function () {
    
        const sheetjs = document.createElement('script');
        sheetjs.lang = "javascript";
        sheetjs.src = "https://cdn.sheetjs.com/xlsx-0.19.2/package/dist/xlsx.full.min.js";
        document.head.appendChild(sheetjs);
      
        const xlsxstyle = document.createElement('script');
        xlsxstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-style@0.8.13/dist/xlsx.core.min.js";
        // document.head.appendChild(xlsxstyle);
      
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
        this._container.prepend(downloadButton);
        downloadButton.addEventListener('click', () => { 
      
          var htmlTable = document.querySelector('table');
          var rows = htmlTable.rows;
          var l = rows.length;
          for (var i = 0; i < l; i++) {
              var cells = rows[i].cells;
              for (var j = 0; j < cells.length; j++) {
                  var cell = cells[j];
              }
          }
      
            var type = "xlsx";
            var tdata = htmlTable;
            var trows = tdata.rows;
            for(var i = 0; i < trows.length; i++){
              var tcells = trows[i].cells;
              for(var j = 0; j < tcells.length; j++){
                var icells = trows[i].cells[j];
                icells.s = {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
              }
            }
      
            const wsheet = XLSX.utils.table_to_sheet(tdata, {origin: 'A4'});
      
            wsheet.A1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.B1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.C1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.D1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.E1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.F1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.G1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.H1 = {v: "C 27.00 - Identification of the counterparty (LE 1)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
           
            wsheet.A2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.B2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.C2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.D2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.E2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.F2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.G2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            wsheet.H2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
            
            wsheet.A4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.B4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.C4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.D4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.E4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.F4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.G4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.H4 = {v: "COUNTERPARTY IDENTIFICATION", t: "s", s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            
            wsheet.A5 = {v: "Code", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.B5 = {v: "Type of Code", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.C5 = {v: "Name", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.D5 = {v: "National Code", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.E5 = {v: "Residence of the Counterparty", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.F5 = {v: "Sector of the Counterparty", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.G5 = {v: "NACE Code", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.H5 = {v: "Type of Counterparty", t: "s", s: {font: {bold: false}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            
            wsheet.A6 = {v: "011", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.B6 = {v: "015", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.C6 = {v: "021", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.D6 = {v: "035", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.E6 = {v: "040", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.F6 = {v: "050", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.G6 = {v: "060", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            wsheet.H6 = {v: "070", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            
      
            for (var x = 6; x < (rows.length + 3); x++){
              for (var y = 0; y < 8; y++){
                const colnamee = XLSX.utils.encode_cell({r:x, c:y});
                const celllval = colnamee;
                wsheet[celllval].s = {alignment: {vertical: "center", horizontal: "center", wrapText: false}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}; 
              }
            }
      
            if(!wsheet["!merges"]) wsheet["!merges"] = [];
            wsheet["!merges"].push(XLSX.utils.decode_range("A1:H1"), XLSX.utils.decode_range("A2:H2"));
            var wbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wbook, wsheet, "C27");
            var wbexport = XLSX.write(wbook, {
                bookType: type,
                bookSST: true,
                type: 'binary',
                cellStyles: true
            }); 
            
            var link = document.createElement("a"); 
            link.download = "target27.xlsx";
            link.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + btoa(wbexport);
            link.click();
            window.open(link, '_blank');
          
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
