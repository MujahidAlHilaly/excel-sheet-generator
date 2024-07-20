let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
tableExists = false

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    if(isNaN(rowsNumber) || rowsNumber<=0 || isNaN(columnsNumber)  || columnsNumber<=0 ){
        swal({
            icon: "error",
            title: "Oops...",
            text: 'Please enter valid row and column numbers!',
        });
        return;
    }
    table.innerHTML = ""
    for(let i=0; i<rowsNumber; i++){
        var tableRow = ""
        for(let j=0; j<columnsNumber; j++){
            tableRow += `<td contenteditable></td>`
        }
        table.innerHTML += tableRow
    }
    if(rowsNumber>0 && columnsNumber>0){
        tableExists = true
    }
}

const ExportToExcel = (type, fn, dl) => {
    if(!tableExists){
        swal({
            icon: "error",
            title: "Oops...",
            text: 'No table to export!',
        });
        return
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}
