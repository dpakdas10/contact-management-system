function addContact() {
  var name = document.getElementById("name").value;
  var phone = document.getElementById("phone").value;
  var email = document.getElementById("email").value;

  var table = document.getElementById("contactTable");

  // Check if contact already exists
  var contactExists = false;
  for (var i = 1; i < table.rows.length; i++) {
    var row = table.rows[i];
    if (
      row.cells[0].textContent === name &&
      row.cells[1].textContent === phone &&
      row.cells[2].textContent === email
    ) {
      contactExists = true;
      break;
    }
  }

  // If contact doesn't exist, add it to the table and clear input fields
  if (!contactExists) {
    var newRow = table.insertRow(table.rows.length);
    var cell1 = newRow.insertCell(0);
    var cell2 = newRow.insertCell(1);
    var cell3 = newRow.insertCell(2);
    var cell4 = newRow.insertCell(3);

    cell1.innerHTML = name;
    cell2.innerHTML = phone;
    cell3.innerHTML = email;
    cell4.innerHTML =
      '<button onclick="deleteContact(this)">Delete</button>' +
      '<button onclick="editContact(this)">Edit</button>';

    // Clear input fields
    document.getElementById("name").value = "";
    document.getElementById("phone").value = "";
    document.getElementById("email").value = "";
  } else {
    alert(
      "Contact with the same name, phone number, and email already exists!"
    );
  }
}

function deleteContact(btn) {
  var row = btn.parentNode.parentNode;
  row.parentNode.removeChild(row);
}

function exportToExcel() {
  var rows = [["Name", "Phone Number", "Email Address"]];
  var table = document.getElementById("contactTable");

  for (var i = 1; i < table.rows.length; i++) {
    var row = table.rows[i];
    var rowData = [
      row.cells[0].textContent,
      row.cells[1].textContent,
      row.cells[2].textContent,
    ];
    rows.push(rowData);
  }

  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Contacts");

  var date = new Date();
  var fileName =
    "contacts_" +
    date.getFullYear() +
    (date.getMonth() + 1) +
    date.getDate() +
    "_" +
    date.getHours() +
    date.getMinutes() +
    date.getSeconds() +
    ".xlsx";

  XLSX.writeFile(wb, fileName);
}

function editContact(btn) {
  var row = btn.parentNode.parentNode;
  var cells = row.cells;

  document.getElementById("editName").value = cells[0].textContent;
  document.getElementById("editPhone").value = cells[1].textContent;
  document.getElementById("editEmail").value = cells[2].textContent;

  // Set up data-index attribute to identify the row to be edited
  document.getElementById("editModal").setAttribute("data-index", row.rowIndex);

  // Show the edit modal
  document.getElementById("editModal").style.display = "flex";
}

function closeModal() {
  // Clear input fields when the modal is closed
  document.getElementById("editName").value = "";
  document.getElementById("editPhone").value = "";
  document.getElementById("editEmail").value = "";

  // Hide the edit modal
  document.getElementById("editModal").style.display = "none";
}

function editExistingContact() {
  var name = document.getElementById("editName").value;
  var phone = document.getElementById("editPhone").value;
  var email = document.getElementById("editEmail").value;

  var table = document.getElementById("contactTable");
  var modal = document.getElementById("editModal");

  // Retrieve the index of the row to be edited from the data-index attribute
  var rowIndex = parseInt(modal.getAttribute("data-index"));

  // Update the table with the edited information
  var row = table.rows[rowIndex];
  var cells = row.cells;

  cells[0].textContent = name;
  cells[1].textContent = phone;
  cells[2].textContent = email;

  closeModal();
}

// Replace the existing saveChanges function with this new one
function saveChanges() {
  var name = document.getElementById("editName").value;
  var phone = document.getElementById("editPhone").value;
  var email = document.getElementById("editEmail").value;

  var table = document.getElementById("contactTable");
  var modal = document.getElementById("editModal");

  // Retrieve the index of the row to be edited from the data-index attribute
  var rowIndex = parseInt(modal.getAttribute("data-index"));

  // Check if the contact already exists
  for (var i = 1; i < table.rows.length; i++) {
    if (i !== rowIndex) {
      var row = table.rows[i];
      var cells = row.cells;

      if (
        cells[0].textContent === name &&
        cells[1].textContent === phone &&
        cells[2].textContent === email
      ) {
        alert("Contact with the same information already exists!");
        return;
      }
    }
  }

  // If the contact doesn't exist, proceed to edit and save changes
  var editedRow = table.rows[rowIndex];
  var editedCells = editedRow.cells;

  editedCells[0].textContent = name;
  editedCells[1].textContent = phone;
  editedCells[2].textContent = email;

  closeModal();
}
