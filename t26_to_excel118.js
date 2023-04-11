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
              background-clip: padding-box;
            }
            .table-cell {
              padding: 5px;
              border-bottom: 1px solid #ccc;
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
            .text-cell {
              mso-number-format: \@;
            }
          </style>
        `;
      // Create a container element to let us center the text.
      const div = document.createElement("div");
    //   div.classList.add('div');
      this._container = element.appendChild(div);
    },
  
    addDownloadButtonListener: function (k) {
      console.log('xyz' + k);

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
      //downloadButton.className = 'download-button';   
      this._container.prepend(downloadButton);
      downloadButton.addEventListener('click', (event) => {

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
    
          wsheet.A1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.C1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.D1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.E1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.F1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.G1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.H1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.I1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.J1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.K1 = {v: "C 26.00 - Large Exposures limits (LE Limits)", t: "s", s: {font: {name: "Calibri", sz: 16, bold: true}, border: {top: {style: "thick"}, left: {style: "thick"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          
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
          
          wsheet.A4 = {v: "", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.A5 = {v: "", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B4 = {v: "", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B5 = {v: "", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          
        //   wsheet["A6"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["A7"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["A8"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["A9"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
          
        //   wsheet["B6"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["B7"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["B8"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
        //   wsheet["B9"].s = {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}};
          
          wsheet.A6 = {v: "010", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.A7 = {v: "020", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.A8 = {v: "030", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.A9 = {v: "040", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          
          wsheet.B6 = {v: "Non institutions", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B7 = {v: "Institutions", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B8 = {v: "Institutions in %", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          wsheet.B9 = {v: "Globally Systemic Important Institutions (G-SIIs)", t: "s", s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
          
          for (var a = 3; a < 5; a++){
            for (var b = 2; b < (k+2); b++){
                const headername = XLSX.utils.encode_cell({r:a, c:b});
                console.log("headername " + headername);
                if(a == 3) 
                wsheet[headername] = {s: {font: {bold: true}, alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium", color: {theme: 4}}}}};
                if(a == 4)
                wsheet[headername] = {s: {alignment: {vertical: "center", horizontal: "center", wrapText: true}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}};
            }
          }
    
          for (var x = 5; x < 9; x++){
            for (var y = 2; y < (k+2); y++){
              const colnamee = XLSX.utils.encode_cell({r:x, c:y});
              const celllval = colnamee;
              console.log("celllval " + celllval);
              wsheet[celllval].s = {alignment: {vertical: "center", horizontal: "center", wrapText: false}, border: {top: {style: "medium"}, left: {style: "medium"}, bottom: {style: "medium"}, right: {style: "medium"}}}; 
            }
          }
          
          // to get range of cells to merge for header
          var headermerge1 = XLSX.utils.encode_range({ s: { c: 2, r: 3 }, e: { c: (k+1), r: 3 } });
          var headermerge2 = XLSX.utils.encode_range({ s: { c: 2, r: 4 }, e: { c: (k+1), r: 4 } });
          console.log(headermerge1 + " and " + headermerge2);

          if(!wsheet["!merges"]) wsheet["!merges"] = [];
          wsheet["!merges"].push(XLSX.utils.decode_range("A1:K1"), XLSX.utils.decode_range("A2:K2"), { s: { c: 2, r: 3 }, e: { c: (k+1), r: 3 } }, { s: { c: 2, r: 4 }, e: { c: (k+1), r: 4 } });
          
          var wbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wbook, wsheet, "C26");
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
            background-clip: padding-box;
          }
          .table-cell {
            padding: 5px;
            border-bottom: 1px solid #ccc;
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
          .table-row {
            border: 1px solid black;
            border-collapse: collapse;
          }
          .text-cell {
              mso-number-format: \@;
            }
    </style>
    `;
      var k = 0;
      for (column_type of ["dimension_like", "measure_like", "table_calculations"]) {
        for (field of queryResponse.fields[column_type]) {
          for (row of data) {
            k++;
          }
          break
        }
      }
      
      console.log('hello.' + k);
        if(k==1){
        generatedHTML += "<p style='font-family:Verdana;align: center;text-align: left;margin-right: auto;margin-left: auto; width:500px;font-weight:bold;font-size:14px;align-items:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 26.00 - Large Exposures limits (LE Limits)</p>";
        generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;margin-left: 55%;text-align: left;padding: 5px;'>* All values reported are in millions </p>";
        }
        else if(k==2){ 
        generatedHTML += "<p style='font-family:Verdana;align: center;text-align: left;margin-right: auto;margin-left: auto; width:600px;font-weight:bold;font-size:14px;align-items:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 26.00 - Large Exposures limits (LE Limits)</p>";
        generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;margin-left:60%;text-align: left;padding: 5px;'>* All values reported are in millions </p>";
        }
        else if(k==3){
        generatedHTML += "<p style='font-family:Verdana;align: center;text-align: left;margin-right: auto;margin-left: auto; width:700px;font-weight:bold;font-size:14px;align-items:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 26.00 - Large Exposures limits (LE Limits)</p>";
        generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;margin-left: 65%;text-align: left;padding: 5px;'>* All values reported are in millions </p>";
        }
        else{ 
        generatedHTML += "<p style='font-family:Verdana;align: center;text-align: left;margin:auto;font-weight:bold;font-size:14px;align-items:left;border:1px solid black;padding: 5px;background-color: #eee;'>C 26.00 - Large Exposures limits (LE Limits)</p>";
        generatedHTML += "<p style='font-family:Verdana;font-size:10px;align-items: center;margin-right: 2%;text-align: right;padding: 5px;'>* All values reported are in millions </p>";
        }
      generatedHTML += `<table class='table'>`;
      generatedHTML += "<tr class='table-header'>";
      generatedHTML += `<th class='table-header' rowspan='2' colspan='2' style='border: 1px solid black;background-color: #eee;color: #eee'></th>`;
      generatedHTML += `<th class='table-header' rowspan='1' colspan='${k}' style='height: 40px;border: 1px solid black;background-color: #eee;font-family: Verdana;'><b>Applicable<br>limit</br></b></th>`;
      generatedHTML += "</tr>";
  
      generatedHTML += "<tr class='table-header'>";
      generatedHTML += `<th class='table-header text-cell' colspan='${k}' style='border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;mso-number-format: "\ \@";'> 010 </th>`;
      generatedHTML += "</tr>";
  
      const header = ['Non institutions', 'Institutions', 'Institutions in %', 'Globally Systemic Important Institutions (G-SIIs)'];
  
      // Loop through the different types of column types looker exposes
      let i = 0;
      const header1 = ['010', '020', '030', '040'];
      for (column_type of ["dimension_like", "measure_like", "table_calculations"]) {
  
        // Look through each field (i.e. row of data)
        for (field of queryResponse.fields[column_type]) {
          // First column is the label
          generatedHTML += `<tr><th class='table-header' style='border: 1px solid black;width:60px;background-color: #eee; padding: 5px;font-family: Verdana;font-weight: normal;mso-number-format: "\ \@";'>${header1[i]}</th>`;
          generatedHTML += `<th class='table-header' style='text-align: left; padding: 5px;width:350px;border: 1px solid black;background-color: #eee;font-family: Verdana;font-weight: normal;'>${header[i]}</th>`;
          // Next columns are the data
          for (row of data) {
            generatedHTML += `<td class='table-cell' style='border: 1px solid black;'>${LookerCharts.Utils.htmlForCell(row[field.name])}</td>`
          }
          generatedHTML += '</tr>';
          i++;
        }
      }
      generatedHTML += "</table>";
      this._container.innerHTML = generatedHTML;
      console.log('abc' + k);
      this.addDownloadButtonListener(k);
  
      done();
    }
  
  });
