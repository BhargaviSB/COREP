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
          margin: auto;
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
        .thead{
          position: sticky;
          top: 0px; 
          z-index: 3;
        }
        th:before {
          content: '';
          top: 0;
          left: 0;
          border-top: 1px solid black;
          position: absolute;
          width: 100%;
      }
       th:after {
        content:''; 
        position:absolute; 
        left: 0; 
        bottom: 0; 
        width:100%; 
        border-bottom: 1px solid rgba(0,0,0,0.12);
      }
      .div{
        overflow-y: auto;
        height: calc(100vh - 100px);
        margin-bottom: 100px;
      }
      </style>
    `;

    // Create a container element to let us center the text.
    const div = document.createElement("div");
    div.classList.add('div');
    this._container = element.appendChild(div);

  },

  addDownloadButtonListener: function () {
    // const cssBoot = document.createElement('link');
    // cssBoot.rel = "stylesheet";
    // cssBoot.href = "https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css";
    // // cssBoot.integrity = "sha384-GLhlTQ8iRABdZLl6O3oVMWSktQOp6b7In1Zl3/Jr59b6EGGoI1aFkw7cmDA6j6gD";
    // cssBoot.crossorigin = "anonymous";
    // document.head.appendChild(cssBoot);
    
    const sheetjs = document.createElement('script');
    sheetjs.lang = "javascript";
    sheetjs.src = "https://cdn.sheetjs.com/xlsx-0.19.2/package/dist/xlsx.full.min.js";
    document.head.appendChild(sheetjs);

    // const fileSaver = document.createElement('script');
    // fileSaver.src = "https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js";
    // document.head.appendChild(fileSaver);

    // const xlsxstyle = document.createElement('script');
    // xlsxstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-style@0.8.13/dist/xlsx.full.min.js";
    // // document.head.appendChild(xlsxstyle);

    const xlsxjsstyle = document.createElement('script');
    // xlsxjsstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js";
    xlsxjsstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
    document.head.appendChild(xlsxjsstyle);

    // const shimm = document.createElement('script');
    // shimm.src = "https://cdn.sheetjs.com/xlsx-latest/package/dist/shim.min.js";
    // document.head.appendChild(shimm);

    // const jsxlsx = document.createElement('script');
    // jsxlsx.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    // document.head.appendChild(jsxlsx);

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
      var l = rows.length;
      for (var i = 0; i < l; i++) {
          var cells = rows[i].cells;
          for (var j = 0; j < cells.length; j++) {
              var cell = cells[j];
              // cells[j].innerHTML = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          }
      }

      // var theader = "<tr class='table-header'><th class='table-header' rowspan='1' colspan='100' style='align-items: left;text-align: left; height: 40px;border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>C 29.00 - Detail of the exposures to individual clients within groups of connected clients (LE 3)</b></th></tr><tr class='table-header'><th class='table-header' rowspan='1' colspan='3' style='background-color:none !important;font-family:Verdana;font-size:10px;align-items: center;text-align: right;padding: 5px;color:grey;font-weight:normal;'>* All values reported are in millions </th></tr>";
      // var theader = document.createElement('h1');
      // var htext = document.createTextNode("C 28.00 - Exposures in the non-trading and trading book (LE 2)");
      // theader.appendChild(htext);

        var type = "xlsx";
        var tdata = htmlTable;
        var wsheet = XLSX.utils.table_to_sheet(tdata, {origin: 'A4'});
        // XLSX.utils.sheet_add_dom(wsheet, theader, {origin: 'A1'});
        // wsheet["!merges"] = [{s:{c:0, r:0}, e:{c:7, r:0}}, {s:{c:0, r:1}, e:{c:7, r:1}}, {s:{c:0, r:3}, e:{c:7, r:3}}];
        wsheet.A1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.B1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.C1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.D1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.E1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.F1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.G1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.H1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.I1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.J1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.K1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.L1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.M1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.N1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.O1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.P1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.Q1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.R1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.S1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.T1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.U1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.V1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.W1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.X1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.Y1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.Z1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AA1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AB1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AC1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AD1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AE1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AF1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AG1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AH1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        wsheet.AI1 = {v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        
        wsheet.A2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.B2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.C2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.D2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.E2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.F2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.G2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.H2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.I2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.J2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.K2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.L2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.M2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.N2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.O2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.P2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.Q2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.R2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.S2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.T2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.U2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.V2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.W2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.X2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.Y2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.Z2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AA2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AB2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AC2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AD2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AE2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AF2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AG2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AH2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};
        wsheet.AI2 = {v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}};

        // wsheet.A4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.D4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.S4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.T4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.U4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.X4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.AF4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        // wsheet.AG4 = {t: "s", s: {border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
        
        
        // wsheet["A1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["B1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["C1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["D1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["E1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["F1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["G1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        // wsheet["H1"].s = {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        if(!wsheet["!merges"]) wsheet["!merges"] = [];
        wsheet["!merges"].push(XLSX.utils.decode_range("A1:AI1"), XLSX.utils.decode_range("A2:AI2"));
        // var sheader = [{v: "C 28.00 - Exposures in the non-trading and trading book (LE 2)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}}];
        // var note = [{v: "* All values reported are in millions", t: "s", s: {font: {name: "Calibri", sz: 9}}}];
        // var range = XLSX.utils.decode_range("A2:AI2");
        // range.s = {fill: {patternType: "solid", fgColor: {rgb: "FFFFFF"}, bgColor: {rgb: "A9AAAB"}}};
        // range.s = {border: {bottom: {style: "medium"}, right: {style: "medium"}}};
        // XLSX.utils.sheet_add_aoa(wsheet, [sheader]);
        // XLSX.utils.sheet_add_aoa(wsheet, [note], { origin: "A2"});
        // range["!merges"] = [XLSX.utils.decode_range("A2:AI2")];
        var wbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbook, wsheet, "C28");
        var wbexport = XLSX.write(wbook, {
            bookType: type,
            bookSST: true,
            type: 'binary',
            cellStyles: true
        }); 
        
        var link = document.createElement("a"); 
        link.download = "target28.xlsx";
        link.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + btoa(wbexport);
        link.click();
        window.open(link, '_blank');
      
    });
},

  // addDownloadButtonListener: function () {
  //   const sheetjs = document.createElement('script');
  //   sheetjs.lang = "javascript";
  //   sheetjs.src = "https://cdn.sheetjs.com/xlsx-0.19.2/package/dist/xlsx.full.min.js";
  //   document.head.appendChild(sheetjs);

  //   const xlsxstyle = document.createElement('script');
  //   xlsxstyle.src = "https://cdn.jsdelivr.net/npm/xlsx-style@0.8.13/dist/xlsx.full.min.js";
  //   document.head.appendChild(xlsxstyle);

  //   const downloadButton = document.createElement('img');
  //   downloadButton.src = "https://cdn.jsdelivr.net/gh/Spoorti-Gandhad/AGBG-Assets@main/downloadAsExcel.jfif";
  //   downloadButton.setAttribute('height', '25px');
  //   downloadButton.setAttribute('width', '25px');
  //   downloadButton.setAttribute('title', 'Download As Excel'); 
  //   downloadButton.style.marginLeft='90%';
  //   this._container.prepend(downloadButton);
  //   downloadButton.addEventListener('click', () => {
  //     var uri = 'data:application/vnd.ms-excel;base64,'
  //       , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{Worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"></head><body><table>{table}</table></body></html>'
  //       , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
  //       , format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) };

  //     // Create a new style element and set the default styles
  //     var table = document.querySelector('table');  
  //     // table.style.type = 'text/css';
  //     // table.style.innerHTML = 'td, th { background-color: white; border: 1px solid black; font-weight: normal; font-size: 11pt; font-family: Calibri; mso-number-format: "\\\@"; }';
  //      var rows = table.rows;
  //     for (var i = 0; i < rows.length; i++) {
  //       var cells = rows[i].cells;
  //       for (var j = 0; j < cells.length; j++) {
  //         var cell = cells[j];
  //       //   cell.setAttribute('style', style);
  //       }
  //     }

  //     // const XLSX = document.createElement('script');
  //     // XLSX.src = 'https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js';
  //     // document.head.appendChild(XLSX);
  //     var ctx = { Worksheet: '28', table: table.innerHTML };
  //     var ctx = { Worksheet: '28', table: "<tr class='table-header'><th class='table-header' rowspan='1' colspan='35' style='align-items: left;text-align: left; height: 40px;border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>C 28.00 - Exposures in the non-trading and trading book (LE 2)</b></th></tr><tr class='table-header'><th class='table-header' rowspan='1' colspan='3' style='background-color:none !important;font-family:Verdana;font-size:10px;align-items: center;text-align: left;padding: 5px;color:grey;font-weight:normal;'>* All values reported are in millions </th></tr>"+table.innerHTML }
          
  //     var xl = format(template, ctx);
  //     const downloadUrl = uri + base64(xl);

  //     // console.log(downloadUrl); // Prints the download URL to the console
  //     //sleep(1000);
  //     window.open(downloadUrl, "_blank");
  //     //const newTab=window.open(downloadUrl, "_blank");
  //   });
  // },


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
          margin: auto;
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
        .thead{
          position: sticky;
          top: 0px; 
          z-index: 3;
        }
        th:before {
          content: '';
          top: 0;
          left: 0;
          border-top: 1px solid black;
          position: absolute;
          width: 100%;
      }
        th:after {
        content:''; 
        position:absolute; 
        left: 0; 
        bottom: 0; 
        width:100%; 
        border-bottom: 1px solid rgba(0,0,0,0.12);
      }
      .div{
        overflow-y: auto;
        height: calc(100vh - 100px);
        margin-bottom: 100px;
      }
      </style>
    `;

    generatedHTML += "<p style='font-family:Verdana;width:4040px;font-weight:bold;font-size:14px;align-items:center;text-align:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 28.00 - Exposures in the non-trading and trading book (LE 2)</p>";
    generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;text-align: right;padding: 5px;'>* All values reported are in millions </p>";
    generatedHTML += "<table class='table'>";
    generatedHTML += "<thead class='thead'>";
    generatedHTML += "<tr class='table-header' >";
    generatedHTML += `<th class='table-header' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>COUNTERPARTY</b><hr style="margin: 0;width: 48.78%;height: 0.6px;top: 27px;position: absolute;left: 0;background-color: black;"></th>`;
    generatedHTML += `<th class='table-header' colspan='15' style='height:25px;border: 1px solid black;background-color: #eee;font-family: Verdana;' ><b>ORIGINAL EXPOSURES</b></th>`;
    generatedHTML += `<th class='table-header' rowspan='5' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Value adjustments and provisions</th>`;
    generatedHTML += `<th class='table-header' rowspan='5' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Exposures deducted from CET 1 or Additional Tier 1 items</th>`;
    generatedHTML += `<th class='table-header' rowspan='3' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Exposure value before application of exemptions and CRM </b><hr style="margin: 0;position: absolute;width: 25.37%;top: 82px;left: 2260px;background-color: black;height: 0.6px;"></th>`;
    generatedHTML += `<th class='table-header' colspan='8' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>ELIGIBLE CREDIT RISK MITIGATION (CRM) TECHNIQUES</b><hr style="height: 0.6px;top: 27px;position: absolute;left: 2603px;width: 23.85%;margin: 0;background-color: black;"></th>`;
    generatedHTML += `<th class='table-header' rowspan='5' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Amounts exempted</th>`;
    generatedHTML += `<th class='table-header' rowspan='3' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Exposure value after application of exemptions and CRM</b><hr style="margin: 0;height: 0.6px;top: 82px;position: absolute;width: 8.4%;left: 3711.5px;background-color: black;"></th>`;
    generatedHTML += "</tr>";
    generatedHTML += "<tr class='table-header'>";
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Code<hr style="margin: 0;height: 0.6px;position: absolute;width: 100%;left: 0;top: 177px;background-color: black;"></th>`;
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Group or individual</th>`;
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Transactions where there is an exposure to underlying assets</th>`;
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Total original exposure</b></th>`;
    generatedHTML += `<th class='table-header' colspan='14' style='height:25px;background-color: #eee;color: #eee'><hr style="background-color: #eee;margin: 0;position: absolute;height: 0.6px;top: 55px;width:37.89%;background-color: #262d33;left: 442px;"></th>`;
    generatedHTML += `<th class='table-header' colspan='6' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Substitution effect of eligible credit risk mitigation techniques</th>`;
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Funded credit protection other than substitution effect</th>`;
    generatedHTML += `<th class='table-header' rowspan='4' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Real estate</th>`;
    generatedHTML += "</tr>";
    generatedHTML += "<tr>";
    generatedHTML += `<th class='table-header' colspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'></th>`;
    generatedHTML += `<th class='table-header' colspan='6' style='height:25px;border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Direct exposures<hr style="margin: 0;position: absolute;height: 0.6px;top: 82.5px;width: 36.16%;left: 442px;background-color: black;"></th>`;
    generatedHTML += `<th class='table-header' colspan='6' style='height:25px;border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Indirect exposures</th>`;
    generatedHTML += `<th class='table-header' rowspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Additional exposures arising from transactions where there is an exposure to underlying assets</th>`;
    generatedHTML += "</tr>";
    generatedHTML += "<tr>";
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'><i>Of which: defaulted</i></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Debt instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Equity instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Derivatives</th>`;
    generatedHTML += `<th class='table-header' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Off balance sheet items<hr style="margin: 0;height: 0.6px;position: absolute;width: 6.52%;top: 134px;background-color: black;left: 956px;"></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Debt instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Equity instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Derivatives</th>`;
    generatedHTML += `<th class='table-header' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Off balance sheet items<hr style="margin: 0;height: 0.6px;position: absolute;width: 6.49%;top: 134px;background-color: black;left: 1644.5px;"></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Total</b></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'><i>Of which: Non-trading book</i></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>% of Tier 1 capital </th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Debt instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Equity instruments</th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Derivatives</th>`;
    generatedHTML += `<th class='table-header' colspan='3' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Off balance sheet items<hr style="margin: 0;height: 0.6px;position: absolute;width: 6.5%;top: 134px;background-color: black;left: 3024.7px;"></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Total</b></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'><i>Of which: Non-trading book</i></th>`;
    generatedHTML += `<th class='table-header' rowspan='2' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>% of Tier 1 capital</th>`;
    generatedHTML += "</tr>";
    generatedHTML += "<tr>";
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Loan commitments</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Financial guarantees</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Other commit-ments</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Loan commitments</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Financial guarantees</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>Other commit-ments</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Loan commitments</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Financial guarantees</th>`;
    generatedHTML += `<th class='table-header' rowspan='1' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>(-) Other commit-ments</th>`;
    generatedHTML += "</tr>";



    const header=['010','020','030','040','050','060','070','080','090','100','110','120','130','140','150','160','170','180','190','200','210','220','230','240','250','260','270','280','290','300','310','320','330','340','350'];
    // First row is the header
    generatedHTML += "<tr class='table-header'>";
    for (let i=0;i<header.length;i++) {
      generatedHTML += `<th class='table-header' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;mso-number-format: "\ \@";'>${header[i]}<hr style="margin: 0;height: 0.6px;position: absolute;width: 100%;left: 0;top: 193px;background-color: black;"></th>`;
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
