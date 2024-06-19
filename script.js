function convertCGPA() {
  

    var cgpa = document.getElementById("cgpa_input").value;
    if(cgpa == ""){
      document.getElementById("error_msg").innerHTML = "Enter valid CGPA (i.e. between 0 & 10)";
      return;
    }
    
    var cgpa = Number(cgpa);
    var grade;
    var per;
    var cls;
    

    if(cgpa < 0 || cgpa > 10){
      document.getElementById("error_msg").innerHTML = "Enter valid CGPA (i.e. between 0 & 10)";
      return;
    }
    else{
      document.getElementById("error_msg").innerHTML = "";
    }

    if(cgpa < 4){
      grade = "F";
      per = "NA";
    }
    else if(cgpa < 4.75){
      grade = "D";
      per = 6.6*cgpa + 13.6;
    }
    else if(cgpa < 5.25){
      grade = "C";
      per = 10*cgpa - 2.5;
    }
    else if(cgpa < 5.75){
      grade = "B";
      per = 10*cgpa - 2.5
    }
    else if(cgpa < 6.75){
      grade = "B+";
      per = 5*cgpa + 26.5;
    }
    else if(cgpa < 8.25){
      grade = "A";
      per = 10*cgpa - 7.5;
    }
    else if(cgpa < 9.5){
      grade = "A+";
      per = 12*cgpa - 25;
    }
    else{
      grade = "O";
      per = 20*cgpa - 100;
    }
    
    if(cgpa < 4)
      cls = "Fail";
    else if(cgpa < 5.5)
      cls = "Pass";
    else if(cgpa < 6.25)
      cls = "Second Class"
    else if(cgpa < 6.75)
      cls = "Higher Second Class";
    else if(cgpa < 7.75)
      cls = "First Class";
    else
      cls = "First Class with Dist";
  

    document.getElementById("op_cgpa").innerHTML = cgpa;
    document.getElementById("op_grade").innerHTML = grade;
    document.getElementById("op_per").innerHTML = per;
    document.getElementById("op_class").innerHTML = cls;
  }
  

  document.getElementById("cgpa_input").addEventListener("keypress", function(event) {
    if (event.keyCode === 13) {
        event.preventDefault();
        document.getElementById("convert_btn").click();
    }
});

function exportTableToCSV(tableID) {
    if (document.getElementById("op_cgpa").innerHTML == '-') {
        document.getElementById("error_msg").innerHTML = "Enter valid CGPA (i.e. between 0 & 10)";
        return;
    }

    var csv = [];
    var rows = document.querySelectorAll(`#${tableID} tr`);

    for (var i = 0; i < rows.length; i++) {
        var row = [], cols = rows[i].querySelectorAll("td, th");

        for (var j = 0; j < cols.length; j++) {
            row.push(cols[j].innerText);
        }

        csv.push(row.join(","));
    }

    var csvFile = new Blob([csv.join("\n")], { type: "text/csv" });

    var downloadLink = document.createElement("a");
    downloadLink.download = "cgpa.csv";
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = "none";

    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

function exportTableToExcel(tableID) {
    if (document.getElementById("op_cgpa").innerHTML == '-') {
        document.getElementById("error_msg").innerHTML = "Enter valid CGPA (i.e. between 0 & 10)";
        return;
    }

    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.table_to_sheet(document.getElementById(tableID));

    XLSX.utils.book_append_sheet(wb, ws, "CGPA Data");

    XLSX.writeFile(wb, "cgpa.xlsx");
}


document.getElementById("download_csv_btn").addEventListener("click", function() {
    exportTableToCSV('opTable');
});

document.getElementById("download_excel_btn").addEventListener("click", function() {
    exportTableToExcel('opTable');
});