"use strict";
document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    var addTimeButton = document.getElementById('add-time');
    var timeFieldsContainer = document.getElementById('time-fields');
    var fileInput = document.getElementById('excel-file');
    var newButton = document.getElementById('new-button');
    var tableBody = document.querySelector('#data-table tbody');
    var timeInputCount = 1;
    var tableData = [];
    addTimeButton.addEventListener('click', function () {
        var timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');
        var newTimeInput = document.createElement('input');
        newTimeInput.type = 'time';
        newTimeInput.name = 'times[]';
        var daysOfWeekContainer = document.createElement('div');
        daysOfWeekContainer.classList.add('days-of-week');
        var days = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'TODOS'];
        days.forEach(function (day) {
            var checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = day;
            checkbox.name = "days-".concat(timeInputCount, "[]");
            daysOfWeekContainer.appendChild(checkbox);
            daysOfWeekContainer.append(day);
        });
        var errorMessage = document.createElement('span');
        errorMessage.classList.add('error-message');
        timeInputContainer.appendChild(newTimeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        timeFieldsContainer.appendChild(timeInputContainer);
        timeInputCount++;
    });
    var validateForm = function (isNewButton) {
        if (isNewButton === void 0) { isNewButton = false; }
        var isValid = true;
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
        var inputs = form.querySelectorAll('input[required], select[required]');
        inputs.forEach(function (input) {
            if (input.id === 'excel-file' && isNewButton)
                return;
            var errorMessageElement = input.nextElementSibling;
            if (!input.value.trim()) {
                isValid = false;
                input.classList.add('error');
                if (errorMessageElement)
                    errorMessageElement.textContent = 'This field is required.';
            }
        });
        var timeInputs = document.querySelectorAll('.time-input');
        timeInputs.forEach(function (timeInput) {
            var timeField = timeInput.querySelector('input[type="time"]');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            var errorMessageElement = timeInput.querySelector('.error-message');
            var timeEntered = timeField.value !== '';
            var daysSelected = checkboxes.length > 0;
            timeField.classList.toggle('error', timeEntered && !daysSelected);
            if (timeEntered !== daysSelected) {
                isValid = false;
                var message = '';
                if (timeEntered && !daysSelected) {
                    message = 'At least one day must be selected.';
                }
                else if (!timeEntered && daysSelected) {
                    message = 'Time is required.';
                }
                errorMessageElement.textContent = message;
            }
        });
        return isValid;
    };
    var renderTable = function () {
        tableBody.innerHTML = '';
        tableData.forEach(function (rowData) {
            var row = document.createElement('tr');
            row.innerHTML = "\n                <td>".concat(rowData['nome'], "</td>\n                <td>").concat(rowData['url'], "</td>\n                <td>").concat(rowData['mensagem'], "</td>\n                <td>").concat(rowData['zoom'], "</td>\n                <td>").concat(rowData['url_teams'], "</td>\n                <td>").concat(rowData['resolucao_tela'], "</td>\n                <td>").concat(rowData['frequencia_tipo'], "</td>\n                <td>").concat(rowData['times'], "</td>\n            ");
            tableBody.appendChild(row);
        });
    };
    newButton.addEventListener('click', function () {
        if (!validateForm(true)) {
            return;
        }
        var formData = new FormData(form);
        var newRow = {};
        formData.forEach(function (value, key) {
            if (key !== 'excel-file' && !key.startsWith('times') && !key.startsWith('days')) {
                newRow[key] = value.toString();
            }
        });
        var timeInputs = document.querySelectorAll('.time-input');
        var formattedTimes = [];
        timeInputs.forEach(function (timeInput) {
            var timeField = timeInput.querySelector('input[type="time"]');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            if (timeField.value && checkboxes.length > 0) {
                var selectedDays_1 = [];
                var hasTodos_1 = false;
                checkboxes.forEach(function (checkbox) {
                    if (checkbox.value === 'TODOS')
                        hasTodos_1 = true;
                    selectedDays_1.push(checkbox.value);
                });
                if (hasTodos_1) {
                    formattedTimes.push("TODOS-".concat(timeField.value));
                }
                else {
                    formattedTimes.push("".concat(selectedDays_1.join('-'), "-").concat(timeField.value));
                }
            }
        });
        newRow['times'] = formattedTimes.join(', ');
        tableData.push(newRow);
        renderTable();
        form.reset();
        // Keep the file input
        var file = fileInput.files ? fileInput.files[0] : null;
        if (file) {
            var dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            fileInput.files = dataTransfer.files;
        }
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        if (!validateForm()) {
            return;
        }
        if (tableData.length === 0) {
            alert('Please add at least one row of data.');
            return;
        }
        var file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            alert('Please upload an Excel file.');
            return;
        }
        var reader = new FileReader();
        reader.onload = function (e) {
            var _a;
            var data = (_a = e.target) === null || _a === void 0 ? void 0 : _a.result;
            var workbook = XLSX.read(data, { type: 'array' });
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
            var dataToAdd = tableData.map(function (row) { return [
                row['nome'],
                row['url'],
                row['mensagem'],
                row['zoom'],
                row['url_teams'],
                row['resolucao_tela'],
                row['frequencia_tipo'],
                row['times']
            ]; });
            XLSX.utils.sheet_add_aoa(worksheet, dataToAdd, { origin: -1 });
            XLSX.writeFile(workbook, file.name);
            alert('The Excel file has been updated and downloaded!');
            form.reset();
            tableData = [];
            renderTable();
        };
        reader.readAsArrayBuffer(file);
    });
});
