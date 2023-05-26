function download_excel(data, headers, sheetName, filename) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    const wsCols = [];

    XLSX.utils.sheet_add_aoa(ws, [headers], {origin: "A1"});

    // ====================== set column width ======================
    data.map((item) => {
        Object.keys(item).map((key, index) => {
            let maxWidth;
            if (typeof item[key] === "number") {
                maxWidth = 10;
            } else if (wsCols[index]) {
                maxWidth = wsCols[index].width < item[key].length ? (item[key].length+5) : wsCols[index].width;
            } else {
                maxWidth = (item[key].length+5)
            }
            wsCols[index] = {width: maxWidth}
        })
    })
    
    ws['!cols'] = wsCols;
    // ==============================================================



    // ====================== set column style ======================
    for (i in ws) {
        if (typeof(ws[i]) != "object") continue;
        let cell = XLSX.utils.decode_cell(i);

        ws[i].s = { // styling for all cells
            font: {
                name: "arial"
            },
            alignment: {
                vertical: "center",
                horizontal: "center",
                wrapText: '1', // any truthy value here
            },
        };

        if (cell.c === 0) { // first column
            ws[i].s.numFmt = "DD/MM/YYYY HH:MM"; // for dates
            ws[i].z = "DD/MM/YYYY HH:MM";
        } else {
            ws[i].s.numFmt = "00.00"; // other numbers
        }

        if (cell.r === 0 ) { // first row
            ws[i].s.fill = { // background color
                patternType: "solid",
                fgColor: { rgb: "b2b2b2" },
                bgColor: { rgb: "b2b2b2" }
            };
        }
    }
    // ==============================================================

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, filename);
}

function read_excel(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'binary'});
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, {raw: true});
        console.log(json);
    }
    reader.readAsBinaryString(file);
}