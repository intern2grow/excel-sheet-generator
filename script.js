let table = document.getElementsByClassName("sheet-body")[0],
     rows = document.getElementsByClassName("rows")[0],
     columns = document.getElementsByClassName("columns")[0];
tableExists = false;

const generateTable = () => {
     let rowsNumber = parseInt(rows.value),
          columnsNumber = parseInt(columns.value);

     if (!rowsNumber || !columnsNumber) {
          Swal.fire({
               icon: "error",
               title: "Oops...",
               text: "Enter a valid number for both inputs",
          });
          return;
     }

     table.innerHTML = "";
     for (let i = 0; i < rowsNumber; i++) {
          var tableRow = "";
          for (let j = 0; j < columnsNumber; j++) {
               tableRow += `<td contenteditable></td>`;
          }
          table.innerHTML += tableRow;
     }
     if (rowsNumber > 0 && columnsNumber > 0) {
          tableExists = true;
     }
};

const ExportToExcel = (type, fn, dl) => {
     if (!tableExists) {
          Swal.fire({
               icon: "error",
               title: "Oops...",
               text: "Generate the table before export!",
          });
          return;
     }

     var elt = table;
     let exportable = true;
     elt.querySelectorAll("td").forEach((it) => {
          if (it.innerText === "") {
               Swal.fire({
                    icon: "error",
                    title: "Oops...",
                    text: "Please fill all table's fields before export",
               });
               exportable = false;
               return;
          }
     });
     if (!exportable) return;
     var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
     console.log("d");
     return dl
          ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
          : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};
