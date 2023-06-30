document.addEventListener("DOMContentLoaded", function() {
  const searchInput = document.getElementById("searchInput");
  const tableBody = document.getElementById("tableBody");
  const searchIcon = document.getElementById("searchIcon");
  const downloadLink = document.getElementById("downloadLink");
  let priceList = [];

  function loadPriceList() {
    const file = "ArchivoCom.xlsx";
    const xhr = new XMLHttpRequest();
    xhr.open("GET", file, true);
    xhr.responseType = "arraybuffer";

    xhr.onload = function(e) {
      const arraybuffer = xhr.response;
      const data = new Uint8Array(arraybuffer);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      priceList = jsonData.slice(1).map(function(row) {
        return {
          codigo: row[0],
          descripcion: row[1],
          precio: parseFloat(row[2]),
          marca: row[3],
          var: row[4]
        };
      });

      displayPriceList(priceList);
    };

    xhr.send();
  }

  function displayPriceList(priceList) {
    tableBody.innerHTML = "";

    priceList.forEach(function(item) {
      const row = document.createElement("tr");
      const codigoCell = createTableCell(item.codigo);
      const descripcionCell = createTableCell(item.descripcion);
      const precioCell = createTableCell(item.precio.toFixed(2));
      const marcaCell = createTableCell(item.marca);
      const varCell = createVarTableCell(item.var);

      row.appendChild(codigoCell);
      row.appendChild(descripcionCell);
      row.appendChild(precioCell);
      row.appendChild(marcaCell);
      row.appendChild(varCell);

      tableBody.appendChild(row);
    });
  }

  function createTableCell(value) {
    const cell = document.createElement("td");
    cell.textContent = value;
    return cell;
  }

  function createVarTableCell(varValue) {
    const cell = document.createElement("td");
    const circle = document.createElement("div");
    circle.className = "circle";

    if (varValue === "I") {
      circle.classList.add("green");
      circle.textContent = "I";
    } else if (varValue === "M") {
      circle.classList.add("red");
      circle.textContent = "M";
    } else if (varValue === "N") {
      circle.classList.add("blue");
      circle.textContent = "N";
    }

    cell.appendChild(circle);
    return cell;
  }

  function searchPriceList() {
    const searchTerm = searchInput.value.toLowerCase();
    const filteredList = filterPriceList(searchTerm);
    displayPriceList(filteredList);
  }

  searchInput.addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
      searchPriceList();
    }
  });

  searchIcon.addEventListener("click", function() {
    searchPriceList();
  });

  function filterPriceList(searchTerm) {
    return priceList.filter(function(item) {
      const lowerCaseTerm = searchTerm.toLowerCase();
      return (
        item.codigo.toLowerCase().includes(lowerCaseTerm) ||
        item.descripcion.toLowerCase().includes(lowerCaseTerm) ||
        item.precio.toFixed(2).includes(lowerCaseTerm) ||
        item.marca.toLowerCase().includes(lowerCaseTerm)
      );
    });
  }

  function exportPriceList() {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(priceList);
    XLSX.utils.book_append_sheet(workbook, worksheet, "ListaPrecios");
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const data = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    if (navigator.msSaveBlob) {
      navigator.msSaveBlob(data, "ListaPrecios.xlsx");
    } else {
      downloadLink.href = window.URL.createObjectURL(data);
    }
  }

  downloadLink.addEventListener("click", function() {
    exportPriceList();
  });

  loadPriceList();
});
