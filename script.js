let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0],
tableExists = false;

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), 
    columnsNumber = parseInt(columns.value);
    table.innerHTML = ""
    for(let i=0; i<rowsNumber; i++){
        var tableRow = "";
        for(let j=0; j<columnsNumber; j++){
            tableRow += `<td contenteditable></td>`;
        }
        table.innerHTML += tableRow;
    }
    if(rowsNumber>0 && columnsNumber>0){
        tableExists = true;
    }
    else{
        Swal.fire({
            icon: 'error',
            title: 'Not Allowed',
            text: 'Please enter the number of rows and columns.',
            confirmButtonText: 'Try again'
        });
    }
}

const ExportToExcel = (type, fn, dl) => {
    if(tableExists){
        var elt = table;
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
        return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')));
    }else {
        Swal.fire({
            icon: 'error',
            title: 'Not Allowed',
            text: 'Please Generate a table before Exporting.',
            confirmButtonText: 'Generate first'
        })
    }
}